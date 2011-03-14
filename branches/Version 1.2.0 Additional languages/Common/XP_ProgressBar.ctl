VERSION 5.00
Begin VB.UserControl XP_ProgressBar 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ScaleHeight     =   990
   ScaleWidth      =   3000
   ToolboxBitmap   =   "XP_ProgressBar.ctx":0000
End
Attribute VB_Name = "XP_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------
'Mario Flores Cool Xp ProgressBar
'Emulating The Windows XP Progress Bar
'Open Source
'6 May 2004
'-----------------------------------------------------------
'Mario Flores Cool Xp ProgressBar 2.0
'MultiStyle ProgressBar
'Open Source
'September 12 2004
'-----------------------------------------------------------

'CD JUAREZ CHIHUAHUA MEXICO

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Integer, ByVal COLORREF As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


'=====================================================
'TEXT FORMAT CONST
Const DT_SINGLELINE   As Long = &H20
Const DT_CALCRECT     As Long = &H400
'=====================================================

'=====================================================
'BORDER FIELD CONST
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
'=====================================================

'=====================================================
'THE POINTAPI STRUCTURE
Private Type POINTAPI
    x As Long                       ' The POINTAPI structure defines the x- and y-coordinates of a point.
    Y As Long
End Type
'=====================================================

'=====================================================
'THE RECT STRUCTURE
Private Type Rect
    Left      As Long     'The RECT structure defines the coordinates of the upper-left and lower-right corners of a rectangle
    Top       As Long
    Right     As Long
    Bottom    As Long
End Type
'=====================================================

'=====================================================
'THE BRUSHSTYLE ENUM
Public Enum BrushStyle
 HS_HORIZONTAL = 0
 HS_VERTICAL = 1
 HS_FDIAGONAL = 2
 HS_BDIAGONAL = 3
 HS_CROSS = 4
 HS_DIAGCROSS = 5
 HS_SOLID = 6
End Enum
'=====================================================

'=====================================================
'THE COOL XP PROGRESSBAR 2.0 STYLES
Public Enum cScrolling
    ccScrollingStandard = 0
    ccScrollingSmooth = 1
    ccScrollingSearch = 2
    ccScrollingOfficeXP = 3
    ccScrollingPastel = 4
    ccScrollingJavT = 5
    ccScrollingMediaPlayer = 6
    ccScrollingCustomBrush = 7
    ccScrollingPicture = 8
    ccScrollingMetallic = 9
End Enum
'=====================================================

'=====================================================
'THE ORIENTATION ENUM
Public Enum cOrientation
    ccOrientationHorizontal = 0
    ccOrientationVertical = 1
End Enum
'=====================================================

'----------------------------------------------------
Private m_Color       As OLE_COLOR
Private m_hDC         As Long
Private m_hWnd        As Long        'PROPERTIES VARIABLES
Private m_Max         As Long
Private m_Min         As Long
Private m_Value       As Long
Private m_ShowText    As Boolean
Private m_Scrolling   As cScrolling
Private m_Orientation As cOrientation
Private m_Brush       As BrushStyle
Private m_Picture     As StdPicture
'----------------------------------------------------

'----------------------------------------------------
Private m_MemDC    As Boolean
Private m_ThDC     As Long
Private m_hBmp     As Long
Private m_hBmpOld  As Long
Private iFnt       As IFont
Private m_fnt      As IFont          'VARIABLES USED IN PROCESS
Private hFntOld    As Long
Private m_lWidth   As Long
Private m_lHeight  As Long
Private fPercent   As Double
Private tR         As Rect
Private TBR        As Rect
Private TSR        As Rect
Private AT         As Rect
Private lSegmentWidth   As Long
Private lSegmentSpacing As Long
'----------------------------------------------------



'==========================================================
'/---Draw ALL ProgressXP Bar  !!!!PUBLIC CALL!!!
'==========================================================

Public Sub DrawProgressBar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020
50030             If m_Value > 100 Then m_Value = 100
50040
50050
50060             GetClientRect m_hWnd, tR               '//--- Reference = Control Client Area
50070
50080             DrawFillRectangle tR, IIf(m_Scrolling = ccScrollingMediaPlayer, &H0, vbWhite), m_hDC '//--- Draw BackGround
50090
50100             '//-- Draw ProgressBar Style
50110
50120             '==========================================================
50130             '/---Draw METALLIC XP STYLE
50140             '==========================================================
50150
50160             If m_Scrolling = ccScrollingMetallic Then
50170
50180                  DrawMetalProgressbar
50190
50200
50210             '==========================================================
50220             '/---Draw OFFICE XP STYLE
50230             '==========================================================
50240
50250             ElseIf m_Scrolling = ccScrollingOfficeXP Then
50260
50270                  DrawOfficeXPProgressbar
50280
50290             '==========================================================
50300             '/---Draw PASTEL XP STYLE
50310             '==========================================================
50320
50330             ElseIf m_Scrolling = ccScrollingPastel Then
50340
50350                  DrawPastelProgressbar
50360
50370             '==========================================================
50380             '/---Draw JAVT XP STYLE
50390             '==========================================================
50400
50410             ElseIf m_Scrolling = ccScrollingJavT Then
50420
50430                  DrawJavTProgressbar
50440
50450             '==========================================================
50460             '/---Draw MEDIA PLAYER XP STYLE
50470             '==========================================================
50480
50490             ElseIf m_Scrolling = ccScrollingMediaPlayer Then
50500
50510                  DrawMediaProgressbar
50520
50530             '==========================================================
50540             '/---Draw CUSTOM BRUSH XP WASH COLOR STYLE
50550             '==========================================================
50560
50570             ElseIf m_Scrolling = ccScrollingCustomBrush Then
50580
50590                  DrawCustomBrushProgressbar
50600
50610             '==========================================================
50620             '/---Draw PICTURE STYLE
50630             '==========================================================
50640
50650             ElseIf m_Scrolling = ccScrollingPicture Then
50660
50670                  DrawPictureProgressbar
50680
50690             Else
50700
50710             '==========================================================
50720             '/---Draw WINDOWS XP STYLE
50730             '==========================================================
50740
50750
50760                 CalcBarSize                            '//--- Calculate Progress and Percent Values
50770
50780                 PBarDraw                               '//--- Draw Scolling Bar (Inside Bar)
50790
50800                 If m_Scrolling = 0 Then DrawDivisions  '//--- Draw SegmentSpacing (This Will Generate the Blocks Effect)
50810
50820                 pDrawBorder                            '//--- Draw The XP Look Border
50830
50840             End If
50850
50860             '==========================================================
50870
50880             DrawTexto                                  '//--- Draw The Percent Text
50890
50900             '==========================================================
50910             '/---Use the AntiFlicker DC
50920             '==========================================================
50930
50940             If m_MemDC Then
50950                 With UserControl
50960                     pDraw .hdc, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
50970                 End With
50980             End If
50990
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawProgressBar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'==========================================================
'/---OFFICE XP STYLE
'==========================================================
Private Sub DrawOfficeXPProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020         DrawRectangle tR, ShiftColorXP(m_Color, 100), m_hDC
50030
50040         With TBR
50050           .Left = 1
50060           .Top = 1
50070           .Bottom = tR.Bottom - 1
50080           .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 100)
50090         End With
50100
50110         DrawFillRectangle TBR, ShiftColorXP(m_Color, 180), m_hDC
50120
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawOfficeXPProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'==========================================================
'/---JAVT XP STYLE
'==========================================================
Private Sub DrawJavTProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020        DrawRectangle tR, ShiftColorXP(m_Color, 10), m_hDC
50030        TBR.Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
50040        DrawGradient m_Color, ShiftColorXP(m_Color, 100), 2, 2, tR.Right - 2, tR.Bottom - 5, m_hDC ', True
50050        DrawGradient ShiftColorXP(m_Color, 250), m_Color, 3, 3, TBR.Right, tR.Bottom - 6, m_hDC  ', True
50060        DrawLine TBR.Right, 2, TBR.Right, tR.Bottom - 2, m_hDC, ShiftColorXP(m_Color, 25)
50070
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawJavTProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'==========================================================
'/---PICTURE STYLE
'==========================================================
Private Sub DrawPictureProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020 Dim Brush      As Long
50030 Dim origBrush  As Long
50040
50050        DrawEdge m_hDC, tR, 2, BF_RECT                       '//--- Draw ProgressBar Border
50060
50070        If Nothing Is m_Picture Then Exit Sub                '//--- In Case No Picture is Choosen
50080
50090        Brush = CreatePatternBrush(m_Picture.handle)         '//-- Use Pattern Picture Draw
50100        origBrush = SelectObject(m_hDC, Brush)
50110        TBR.Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
50120
50130        PatBlt m_hDC, 2, 2, TBR.Right, tR.Bottom - 4, vbPatCopy
50140
50150        SelectObject m_hDC, origBrush
50160        DeleteObject Brush
50170
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawPictureProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'==========================================================
'/---PASTEL XP STYLE
'==========================================================
Private Sub DrawPastelProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010         DrawEdge m_hDC, tR, 6, BF_RECT
50020         DrawGradient ShiftColorXP(m_Color, 140), ShiftColorXP(m_Color, 200), 2, 2, tR.Left + (tR.Right - tR.Left - 4) * (m_Value / 100), tR.Bottom - 3, m_hDC, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawPastelProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'==========================================================
'/---METALLIC XP STYLE
'==========================================================
Private Sub DrawMetalProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010          TBR.Right = tR.Left + (tR.Right - tR.Left - 4) * (m_Value / 100)
50020
50030          DrawGradient vbWhite, &HC0C0C0, 2, 2, tR.Right - 3, (tR.Bottom - 3) / 2, m_hDC
50040          DrawGradient BlendColor(&HC0C0C0, &H0, 255), &HC0C0C0, 2, (tR.Bottom - 3) / 2, tR.Right - 3, (tR.Bottom - 3) / 2, m_hDC
50050          DrawGradient ShiftColorXP(m_Color, 150), BlendColor(m_Color, &H0, 180), 2, 2, TBR.Right, (tR.Bottom - 3) / 2, m_hDC
50060          DrawGradient BlendColor(m_Color, &H0, 190), m_Color, 2, (tR.Bottom - 3) / 2, TBR.Right, (tR.Bottom - 3) / 2, m_hDC
50070
50080          tR.Left = tR.Left + 3
50090          pDrawBorder
50100
50110
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawMetalProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'==========================================================
'/---CUSTOM BRUSH XP STYLE
'==========================================================
Private Sub DrawCustomBrushProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020    Dim hBrush As Long
50030
50040    DrawEdge m_hDC, tR, 9, BF_RECT
50050
50060    With TBR
50070       .Left = 2
50080       .Top = 2
50090       .Bottom = tR.Bottom - 2
50100       .Right = tR.Left + (tR.Right - tR.Left) * (m_Value / 101)
50110    End With
50120
50130    hBrush = CreateHatchBrush(m_Brush, GetLngColor(Color))
50140    SetBkColor m_hDC, ShiftColorXP(m_Color, 140)
50150    FillRect m_hDC, TBR, hBrush
50160    DeleteObject hBrush
50170
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawCustomBrushProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'==========================================================
'/---MEDIA PROGRESS XP STYLE
'==========================================================
Private Sub DrawMediaProgressbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020         DrawRectangle tR, BlendColor(m_Color, &H0, 200), m_hDC
50030         DrawGradient &H0&, ShiftColorXP(GetLngColor(BlendColor(m_Color, &H0, 100)), 10), 2, 2, tR.Left + (tR.Right - tR.Left - 5) * (m_Value / 100), tR.Bottom - 2, m_hDC, True
50040
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawMediaProgressbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'==========================================================
'/---Calculate Division Bars & Percent Values
'==========================================================

Private Sub CalcBarSize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020       lSegmentWidth = IIf(m_Scrolling = 0, 6, 0) '/-- Windows Default
50030       lSegmentSpacing = 2                        '/-- Windows Default
50040
50050       tR.Left = tR.Left + 3
50060
50070       LSet TBR = tR
50080
50090       fPercent = m_Value / 98
50100
50110       If fPercent < 0# Then fPercent = 0#
50120
50130       If m_Orientation = 0 Then
50140
50150       '=======================================================================================
50160       '                                 Calc Horizontal ProgressBar
50170       '---------------------------------------------------------------------------------------
50180
50190          TBR.Right = tR.Left + (tR.Right - tR.Left) * fPercent
50200
50210          TBR.Right = TBR.Right - ((TBR.Right - TBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
50220
50230          If TBR.Right < tR.Left Then
50240             TBR.Right = tR.Left
50250          End If
50260
50270       Else
50280
50290       '=======================================================================================
50300       '                                 Calc Vertical ProgressBar
50310       '---------------------------------------------------------------------------------------
50320          fPercent = 1# - fPercent
50330          TBR.Top = tR.Top + (tR.Bottom - tR.Top) * fPercent
50340          TBR.Top = TBR.Top - ((TBR.Top - TBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
50350          If TBR.Top > tR.Bottom Then TBR.Top = tR.Bottom
50360
50370
50380
50390       End If
50400
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "CalcBarSize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'==========================================================
'/---Draw Division Bars
'==========================================================

Private Sub DrawDivisions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  Dim hBR As Long
50030
50040   hBR = CreateSolidBrush(vbWhite)
50050
50060       LSet TSR = tR
50070
50080
50090       If m_Orientation = 0 Then
50100
50110
50120       '=======================================================================================
50130       '                                 Draw Horizontal ProgressBar
50140       '---------------------------------------------------------------------------------------
50150          For i = TBR.Left + lSegmentWidth To TBR.Right Step lSegmentWidth + lSegmentSpacing
50160             TSR.Left = i + 1
50170             TSR.Right = i + 1 + lSegmentSpacing
50180             FillRect m_hDC, TSR, hBR
50190          Next i
50200       '---------------------------------------------------------------------------------------
50210
50220       Else
50230
50240       '=======================================================================================
50250       '                                  Draw Vertical ProgressBar
50260       '---------------------------------------------------------------------------------------
50270          For i = TBR.Bottom To TBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
50280             TSR.Top = i - 2
50290             TSR.Bottom = i - 2 + lSegmentSpacing
50300             FillRect m_hDC, TSR, hBR
50310          Next i
50320        '---------------------------------------------------------------------------------------
50330
50340       End If
50350
50360       DeleteObject hBR
50370
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawDivisions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


'==========================================================
'/---Draw The ProgressXP Bar Border  ;)
'==========================================================

Private Sub pDrawBorder()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim RTemp As Rect
50020
50030  tR.Left = tR.Left - 3
50040
50050  Let RTemp = tR
50060
50070
50080  DrawLine 2, 1, tR.Right - 2, 1, m_hDC, &HBEBEBE
50090  DrawLine 2, tR.Bottom - 2, tR.Right - 2, tR.Bottom - 2, m_hDC, &HEFEFEF
50100  DrawLine 1, 2, 1, tR.Bottom - 2, m_hDC, &HBEBEBE
50110  DrawLine 2, 2, 2, tR.Bottom - 2, m_hDC, &HEFEFEF
50120  DrawLine 2, 2, tR.Right - 2, 2, m_hDC, &HEFEFEF
50130  DrawLine tR.Right - 2, 2, tR.Right - 2, tR.Bottom - 2, m_hDC, &HEFEFEF
50140
50150  DrawRectangle tR, GetLngColor(&H686868), m_hDC
50160
50170
50180  Call SetPixelV(m_hDC, 0, 0, GetLngColor(vbWhite))
50190  Call SetPixelV(m_hDC, 0, 1, GetLngColor(&HA6ABAC))
50200  Call SetPixelV(m_hDC, 0, 2, GetLngColor(&H7D7E7F))
50210  Call SetPixelV(m_hDC, 1, 0, GetLngColor(&HA7ABAC)) '//TOP RIGHT CORNER
50220  Call SetPixelV(m_hDC, 1, 1, GetLngColor(&H777777))
50230  Call SetPixelV(m_hDC, 2, 0, GetLngColor(&H7D7E7F))
50240  Call SetPixelV(m_hDC, 2, 2, GetLngColor(&HBEBEBE))
50250
50260  Call SetPixelV(m_hDC, 0, tR.Bottom - 1, GetLngColor(vbWhite))
50270  Call SetPixelV(m_hDC, 1, tR.Bottom - 1, GetLngColor(&HA6ABAC))
50280  Call SetPixelV(m_hDC, 2, tR.Bottom - 1, GetLngColor(&H7D7E7F))
50290  Call SetPixelV(m_hDC, 0, tR.Bottom - 3, GetLngColor(&H7D7E7F)) '//BOTTOM RIGHT CORNER
50300  Call SetPixelV(m_hDC, 0, tR.Bottom - 2, GetLngColor(&HA7ABAC))
50310  Call SetPixelV(m_hDC, 1, tR.Bottom - 2, GetLngColor(&H777777))
50320
50330  Call SetPixelV(m_hDC, tR.Right - 1, 0, GetLngColor(vbWhite))
50340  Call SetPixelV(m_hDC, tR.Right - 1, 1, GetLngColor(&HBEBEBE))
50350  Call SetPixelV(m_hDC, tR.Right - 1, 2, GetLngColor(&H7D7E7F)) '//TOP LEFT CORNER
50360  Call SetPixelV(m_hDC, tR.Right - 2, 2, GetLngColor(&HBEBEBE))
50370  Call SetPixelV(m_hDC, tR.Right - 2, 1, GetLngColor(&H686868))
50380
50390  Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 1, GetLngColor(vbWhite))
50400  Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 2, GetLngColor(&HBEBEBE))
50410  Call SetPixelV(m_hDC, tR.Right - 1, tR.Bottom - 3, GetLngColor(&H7D7E7F))
50420  Call SetPixelV(m_hDC, tR.Right - 2, tR.Bottom - 2, GetLngColor(&H777777)) '//TOP RIGHT CORNER
50430  Call SetPixelV(m_hDC, tR.Right - 2, tR.Bottom - 1, GetLngColor(&HBEBEBE))
50440  Call SetPixelV(m_hDC, tR.Right - 3, tR.Bottom - 1, GetLngColor(&H7D7E7F))
50450
50460
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "pDrawBorder")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


'==========================================================
'/---Draw The ProgressXP Bar ;)
'==========================================================

Private Sub PBarDraw()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim TempRect As Rect
50020 Dim ITemp    As Long
50030
50040 If m_Orientation = 0 Then
50050
50060     If TBR.Right <= 14 Then TBR.Right = 12
50070
50080     TempRect.Left = 4
50090     TempRect.Right = IIf(TBR.Right + 4 > tR.Right, TBR.Right - 4, TBR.Right)
50100     TempRect.Top = 8
50110     TempRect.Bottom = tR.Bottom - 8
50120
50130     '=======================================================================================
50140     '                                 Draw Horizontal ProgressBar
50150     '---------------------------------------------------------------------------------------
50160
50170
50180      If m_Scrolling = ccScrollingSearch Then
50190          GoSub HorizontalSearch
50200      Else
50210         DrawGradient ShiftColorXP(m_Color, 150), m_Color, 4, 3, TempRect.Right, 6, m_hDC
50220         DrawFillRectangle TempRect, m_Color, m_hDC
50230         DrawGradient m_Color, ShiftColorXP(m_Color, 150), 4, TempRect.Bottom - 2, TempRect.Right, 6, m_hDC
50240      End If
50250 Else
50260
50270     TempRect.Left = 9
50280     TempRect.Right = tR.Right - 8
50290     TempRect.Top = TBR.Top
50300     TempRect.Bottom = tR.Bottom
50310
50320     '=======================================================================================
50330     '                                 Draw Vertical ProgressBar
50340     '---------------------------------------------------------------------------------------
50350
50360     If m_Scrolling = ccScrollingSearch Then
50370          GoSub VerticalSearch
50380     Else
50390         DrawGradient ShiftColorXP(m_Color, 150), m_Color, 4, TBR.Top, 4, tR.Bottom, m_hDC, True
50400         DrawFillRectangle TempRect, m_Color, m_hDC
50410         DrawGradient m_Color, ShiftColorXP(m_Color, 150), tR.Right - 8, TBR.Top, 4, tR.Bottom, m_hDC, True
50420     End If
50430
50440     '--------------------   <-------- Gradient Color From (- to +)
50450     '||||||||||||||||||||   <-------- Fill Color
50460     '--------------------   <-------- Gradient Color From (+ to -)
50470
50480 End If
50490
50500 Exit Sub
50510
HorizontalSearch:
50530
50540
50550     For ITemp = 0 To 2
50560
50570         With TempRect
50580           .Left = TBR.Right + ((lSegmentSpacing + 10) * (ITemp)) - (45 * ((100 - m_Value) / 100))
50590           .Right = .Left + 10
50600           .Top = 8
50610           .Bottom = tR.Bottom - 8
50620           DrawGradient ShiftColorXP(m_Color, 220 - (40 * ITemp)), ShiftColorXP(m_Color, 200 - (40 * ITemp)), .Left, 3, 9, tR.Bottom - 2, m_hDC, True
50630         End With
50640
50650     Next ITemp
50660
50670 Return
50680
VerticalSearch:
50700
50710
50720     For ITemp = 0 To 2
50730
50740         With TempRect
50750           .Left = 8
50760           .Right = tR.Right - 8
50770           .Top = TBR.Top + ((lSegmentSpacing + 10) * ITemp)
50780           .Bottom = .Top + 10
50790           DrawGradient ShiftColorXP(m_Color, 220 - (40 * ITemp)), ShiftColorXP(m_Color, 200 - (40 * ITemp)), tR.Right - 2, .Top, 2, 9, m_hDC
50800         End With
50810
50820     Next ITemp
50830
50840 Return
50850
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "PBarDraw")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'======================================================================
'DRAWS THE PERCENT TEXT ON PROGRESS BAR
Private Function DrawTexto()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim ThisText As String
50020 Dim isAlpha  As Boolean
50030
50040 If (m_Scrolling = ccScrollingMediaPlayer Or m_Scrolling = ccScrollingMetallic) Then isAlpha = True
50050
50060
50070  If m_Scrolling = ccScrollingSearch Then
50080     ThisText = "Searching.."
50090  Else
50100     ThisText = Round(m_Value) & " %"
50110  End If
50120
50130  If (m_ShowText) Then
50140
50150       Set iFnt = Font                             '//--New Font
50160       hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
50170       SetBkMode m_hDC, 1                          '//--Transparent Text
50180
50190       '//--Use the Alpha Text Color Look if Progress is MediaPlayer Style, else Normal (Gray)
50200       SetTextColor m_hDC, GetLngColor(IIf(m_Scrolling = ccScrollingMediaPlayer, &HC0C0C0, vbBlack))
50210
50220       CalculateAlphaTextRect ThisText             '//--Calculate The Text Rectangle
50230
50240       '//-- If ProgressBar is already over the Text don't draw the old text, yust draw the Alpha Text
50250            'It saves some memory
50260
50270       If ((tR.Right * (m_Value / 100)) <= AT.Right) Or Not isAlpha Then
50280 '50280             DrawText m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE
50290             DrawText m_hDC, ThisText, -1, AT, DT_SINGLELINE
50300       End If
50310
50320       SelectObject m_hDC, hFntOld  'Delete the Used Font
50330
50340       '//--Use the Alpha Text Look if Progress is AlPhA Style
50350       If isAlpha Then DrawAlphaText ThisText
50360
50370  End If
50380
50390
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawTexto")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
'======================================================================

'======================================================================
'ALPHA TEXT RECT FUNCTION
Private Sub CalculateAlphaTextRect(ByVal ThisText As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020       '//--Calculates the Bounding Rects Of the Text using DT_CALCRECT
50030 '50030       DrawText m_hDC, ThisText, Len(ThisText), AT, DT_CALCRECT
50040       DrawText m_hDC, ThisText, -1, AT, DT_CALCRECT
50050       AT.Left = (tR.Right / 2) - ((AT.Right - AT.Left) / 2)
50060       AT.Top = (tR.Bottom / 2) - ((AT.Bottom - AT.Top) / 2)
50070
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "CalculateAlphaTextRect")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'ALPHA TEXT FUNCTION
Private Sub DrawAlphaText(ByVal ThisText As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Set iFnt = Font                             '//--New Font
50030  hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
50040  SetBkMode m_hDC, 1                          '//--Transparent Text
50050
50060
50070         '//-- This is When the Text is Drawn
50080             '//--Gives the Media Player Text Look (Changes Color When Progress is over the Text)
50090
50100             If (tR.Right * (m_Value / 100)) >= AT.Left Then
50110                 SetTextColor m_hDC, GetLngColor(IIf(m_Scrolling = ccScrollingMediaPlayer, ShiftColorXP(m_Color, 80), vbWhite))
50120                 AT.Left = (tR.Right / 2) - ((AT.Right - AT.Left) / 2)
50130                 AT.Right = (tR.Right * (m_Value / 100))
50140                 DrawText m_hDC, ThisText, Len(ThisText), AT, DT_SINGLELINE
50150                 SelectObject m_hDC, hFntOld
50160             End If
50170
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawAlphaText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     If (Color And &H80000000) Then
50030         GetLngColor = GetSysColor(Color And &H7FFFFFFF)
50040     Else
50050         GetLngColor = Color
50060     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "GetLngColor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
'======================================================================

'======================================================================
'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawRectangle(ByRef BRect As Rect, ByVal Color As Long, ByVal hdc As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020 Dim hBrush As Long
50030
50040     hBrush = CreateSolidBrush(Color)
50050     FrameRect hdc, BRect, hBrush
50060     DeleteObject hBrush
50070
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawRectangle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'DRAWS A LINE WITH A DEFINED COLOR
Public Sub DrawLine( _
           ByVal x As Long, _
           ByVal Y As Long, _
           ByVal Width As Long, _
           ByVal Height As Long, _
           ByVal cHdc As Long, _
           ByVal Color As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     Dim Pen1    As Long
50030     Dim Pen2    As Long
50040     Dim Outline As Long
50050     Dim pos     As POINTAPI
50060
50070     Pen1 = CreatePen(0, 1, GetLngColor(Color))
50080     Pen2 = SelectObject(cHdc, Pen1)
50090
50100         MoveToEx cHdc, x, Y, pos
50110         LineTo cHdc, Width, Height
50120
50130     SelectObject cHdc, Pen2
50140     DeleteObject Pen2
50150     DeleteObject Pen1
50160
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawLine")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'BLENDS AN SPECIFIED COLOR TO GET XP COLOR LOOK
Private Function ShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     Dim R As Long, G As Long, b As Long, Delta As Long
50030
50040     R = (MyColor And &HFF)
50050     G = ((MyColor \ &H100) Mod &H100)
50060     b = ((MyColor \ &H10000) Mod &H100)
50070
50080     Delta = &HFF - Base
50090
50100     b = Base + b * Delta \ &HFF
50110     G = Base + G * Delta \ &HFF
50120     R = Base + R * Delta \ &HFF
50130
50140     If R > 255 Then R = 255
50150     If G > 255 Then G = 255
50160     If b > 255 Then b = 255
50170
50180     ShiftColorXP = R + 256& * G + 65536 * b
50190
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "ShiftColorXP")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
'======================================================================

'======================================================================
'DRAWS A 2 COLOR GRADIENT AREA WITH A PREDEFINED DIRECTION
Public Sub DrawGradient(lEndColor As Long, lStartcolor As Long, ByVal x As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal hdc As Long, Optional bH As Boolean)
    On Error Resume Next
    
    ''Draw a Vertical Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    sR = (sR - eR) / IIf(bH, X2, Y2)
    sG = (sG - eG) / IIf(bH, X2, Y2)
    sB = (sB - eB) / IIf(bH, X2, Y2)
    
        
    For ni = 0 To IIf(bH, X2, Y2)
        
        If bH Then
            DrawLine x + ni, Y, x + ni, Y2, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        Else
            DrawLine x, Y + ni, X2, Y + ni, hdc, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
        End If
        
    Next ni
End Sub
'======================================================================

'======================================================================
'BLENDS 2 COLORS WITH A PREDEFINED ALPHA VALUE
Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim lCFrom As Long
50020 Dim lCTo As Long
50030 Dim lSrcR As Long
50040 Dim lSrcG As Long
50050 Dim lSrcB As Long
50060 Dim lDstR As Long
50070 Dim lDstG As Long
50080 Dim lDstB As Long
50090
50100    lCFrom = GetLngColor(oColorFrom)
50110    lCTo = GetLngColor(oColorTo)
50120
50130    lSrcR = lCFrom And &HFF
50140    lSrcG = (lCFrom And &HFF00&) \ &H100&
50150    lSrcB = (lCFrom And &HFF0000) \ &H10000
50160    lDstR = lCTo And &HFF
50170    lDstG = (lCTo And &HFF00&) \ &H100&
50180    lDstB = (lCTo And &HFF0000) \ &H10000
50190
50200    BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
      )
50250
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "BlendColor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
'======================================================================

'======================================================================
'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawFillRectangle(ByRef hRect As Rect, ByVal Color As Long, ByVal MyHdc As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020 Dim hBrush As Long
50030
50040    hBrush = CreateSolidBrush(GetLngColor(Color))
50050    FillRect MyHdc, hRect, hBrush
50060    DeleteObject hBrush
50070
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "DrawFillRectangle")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
Private Function ThDC(Width As Long, Height As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    If m_ThDC = 0 Then
50020       If (Width > 0) And (Height > 0) Then
50030          pCreate Width, Height
50040       End If
50050    Else
50060       If Width > m_lWidth Or Height > m_lHeight Then
50070          pCreate Width, Height
50080       End If
50090    End If
50100    ThDC = m_ThDC
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "ThDC")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
'======================================================================

'======================================================================
'CREATES THE TEMP DC
Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim lhDCC As Long
50020    pDestroy
50030    lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
50040    If Not (lhDCC = 0) Then
50050       m_ThDC = CreateCompatibleDC(lhDCC)
50060       If Not (m_ThDC = 0) Then
50070          m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
50080          If Not (m_hBmp = 0) Then
50090             m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
50100             If Not (m_hBmpOld = 0) Then
50110                m_lWidth = Width
50120                m_lHeight = Height
50130                DeleteDC lhDCC
50140                Exit Sub
50150             End If
50160          End If
50170       End If
50180       DeleteDC lhDCC
50190       pDestroy
50200    End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "pCreate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'DRAWS THE TEMP DC
Public Sub pDraw( _
      ByVal hdc As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    If WidthSrc <= 0 Then WidthSrc = m_lWidth
50020    If HeightSrc <= 0 Then HeightSrc = m_lHeight
50030    BitBlt hdc, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy
50040
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "pDraw")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================

'======================================================================
'DESTROYS THE TEMP DC
Private Sub pDestroy()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    If Not m_hBmpOld = 0 Then
50020       SelectObject m_ThDC, m_hBmpOld
50030       m_hBmpOld = 0
50040    End If
50050    If Not m_hBmp = 0 Then
50060       DeleteObject m_hBmp
50070       m_hBmp = 0
50080    End If
50090    If Not m_ThDC = 0 Then
50100       DeleteDC m_ThDC
50110       m_ThDC = 0
50120    End If
50130    m_lWidth = 0
50140    m_lHeight = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "pDestroy")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'======================================================================



'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================
'USER CONTROL EVENTS
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================


Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020      Dim fnt As New StdFont
50030          Set Font = fnt
50040
50050      With UserControl
50060         .BackColor = vbWhite
50070         .ScaleMode = vbPixels
50080      End With
50090
50100      '----------------------------------------------------------
50110      'Default Values
50120      hdc = UserControl.hdc
50130      hwnd = UserControl.hwnd
50140      m_Max = 100
50150      m_Min = 0
50160      m_Value = 0
50170      m_Orientation = ccOrientationHorizontal
50180      m_Scrolling = ccScrollingStandard
50190      m_Color = GetLngColor(vbHighlight)
50200      DrawProgressBar
50210      '----------------------------------------------------------
50220
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Paint()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "UserControl_Paint")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 hdc = UserControl.hdc
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Terminate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  pDestroy 'Destroy Temp DC
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "UserControl_Terminate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================
'USER CONTROL PROPERTIES
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'===========================================================================

Public Property Let BrushStyle(ByVal Style As BrushStyle)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_Brush = Style
50020    PropertyChanged "BrushStyle"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "BrushStyle [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/sets the color of the ProgressBar"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Color = m_Color
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Color [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Color(ByVal lColor As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_Color = GetLngColor(lColor)
50020    DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Color [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Font() As IFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Set Font = m_fnt
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Font [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Set Font(ByRef fnt As IFont)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Set m_fnt = fnt    'Defined By System but can change by user choice.(ADD Property!!)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Font [SET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Font(ByRef fnt As IFont)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Set m_fnt = fnt
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Font [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get hwnd() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    hwnd = m_hWnd
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "hwnd [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let hwnd(ByVal chWnd As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_hWnd = chWnd
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "hwnd [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get hdc() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    hdc = m_hDC
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "hdc [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let hdc(ByVal cHdc As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010      '=============================================
50020    'AntiFlick...Cleaner HDC
50030    m_hDC = ThDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
50040
50050    If m_hDC = 0 Then
50060       m_hDC = UserControl.hdc   'On Fail...Do it Normally
50070    Else
50080       m_MemDC = True
50090    End If
50100    '=============================================
50110
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "hdc [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Image() As StdPicture
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     If Nothing Is m_Picture Then Exit Property
50020     Set Image = m_Picture
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Image [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Set Image(ByVal handle As StdPicture)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Set m_Picture = handle
50020    PropertyChanged "Image"
50030    DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Image [SET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Min() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Min = m_Min
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Min [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Min(ByVal cMin As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_Min = cMin
50020    PropertyChanged "Min"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Min [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Max() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Max = m_Max
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Max [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Max(ByVal cMax As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_Max = cMax
50020    PropertyChanged "Max"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Max [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Orientation() As cOrientation
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Orientation = m_Orientation
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Orientation [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Orientation(ByVal cOrientation As cOrientation)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_Orientation = cOrientation
50020    PropertyChanged "Orientation"
50030    DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Orientation [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Scrolling() As cScrolling
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    Scrolling = m_Scrolling
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Scrolling [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Scrolling(ByVal lScrolling As cScrolling)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_Scrolling = lScrolling
50020    PropertyChanged "Scrolling"
50030    DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "Scrolling [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ShowText() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    ShowText = m_ShowText
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "ShowText [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let ShowText(ByVal bShowText As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    m_ShowText = bShowText
50020    PropertyChanged "ShowText"
50030    DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "ShowText [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get value() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010    value = ((m_Value / 100) * m_Max) / IIf(m_Min > 0, m_Min, 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "value [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let value(ByVal cValue As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_Value = ((cValue * 100) / m_Max) + m_Min
50020     'PropertyChanged "Value"
50030     DrawProgressBar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "value [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

'=======================================================================================================================
' USERCONTROL WRITE PROPERTIES
'=======================================================================================================================

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Call PropBag.WriteProperty("Font", Font)
50020  Call PropBag.WriteProperty("BrushStyle", m_Brush, 4)
50030  Call PropBag.WriteProperty("Color", m_Color, vbHighlight)
50040  Call PropBag.WriteProperty("Image", m_Picture, Nothing)
50050  Call PropBag.WriteProperty("Max", m_Max, 100)
50060  Call PropBag.WriteProperty("Min", m_Min, 0)
50070  Call PropBag.WriteProperty("Orientation", m_Orientation, ccOrientationHorizontal)
50080  Call PropBag.WriteProperty("Scrolling", m_Scrolling, ccScrollingStandard)
50090  Call PropBag.WriteProperty("ShowText", m_ShowText, False)
50100  Call PropBag.WriteProperty("Value", m_Value, 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "UserControl_WriteProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
 End Sub

'=======================================================================================================================
' USERCONTROL READ PROPERTIES
'=======================================================================================================================

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Set Font = PropBag.ReadProperty("Font")
50020 m_Brush = PropBag.ReadProperty("BrushStyle", 4)
50030 Color = PropBag.ReadProperty("Color", vbHighlight)
50040 Set m_Picture = PropBag.ReadProperty("Image", Nothing)
50050 Max = PropBag.ReadProperty("Max", 100)
50060 Min = PropBag.ReadProperty("Min", 0)
50070 Orientation = PropBag.ReadProperty("Orientation", ccOrientationHorizontal)
50080 Scrolling = PropBag.ReadProperty("Scrolling", ccScrollingStandard)
50090 ShowText = PropBag.ReadProperty("ShowText", False)
50100 value = PropBag.ReadProperty("Value", 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("XP_ProgressBar", "UserControl_ReadProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

