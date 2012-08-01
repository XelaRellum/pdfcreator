VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' http://www.codeproject.com/Articles/7945/How-to-implement-Windows-XP-style-balloon-tooltip
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

''Windows API Functions
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Long, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As Long, _
        ByVal x As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hWndParent As Long, _
        ByVal hMenu As Long, _
        ByVal hInstance As Long, _
        ByRef lpParam As Any) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByRef lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Private Declare Function DestroyWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long

Private Declare Function ClientToScreen Lib "user32" ( _
        ByVal hwnd As Long, _
        lpPoint As POINTAPI) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

'Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTM_TRACKACTIVATE = WM_USER + 17
Private Const TTM_TRACKPOSITION = WM_USER + 18
Private Const WS_POPUP = &H80000000

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As Rect
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'local variable(s) to hold property value(s)
Private mvarBackColor As Long
Private mvarTitle As String
Private mvarForeColor As Long
Private mvarIcon As ttIconType
Private mvarCentered As Boolean
Private mvarStyle As ttStyleEnum
Private mvarTipText As String
Private mvarVisibleTime As Long
Private mvarDelayTime As Long
Private mvarPopupOnDemand As Boolean

'private data
Private m_lTTHwnd As Long                                  ' hwnd of the tooltip
Private m_lParentHwnd As Long                              ' hwnd of the window the tooltip attached to
Private ti As TOOLINFO

Private Sub Class_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     InitCommonControls
50020     mvarDelayTime = 500
50030     mvarVisibleTime = 5000
50040     mvarPopupOnDemand = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Class_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
Private Sub Class_Terminate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Destroy
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Class_Terminate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
'//////////////////////////////////////////////////////
Public Property Get VisibleTime() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     VisibleTime = mvarVisibleTime
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "VisibleTime [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Let VisibleTime(ByVal lData As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarVisibleTime = lData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "VisibleTime [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Get DelayTime() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     DelayTime = mvarDelayTime
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "DelayTime [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Let DelayTime(ByVal lData As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarDelayTime = lData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "DelayTime [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let Icon(ByVal vData As ttIconType)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarIcon = vData
50020     If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
50030         SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
50040     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Icon [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get Icon() As ttIconType
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Icon = mvarIcon
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Icon [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let ForeColor(ByVal vData As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarForeColor = vData
50020     If m_lTTHwnd <> 0 Then
50030         SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
50040     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "ForeColor [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get ForeColor() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     ForeColor = mvarForeColor
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "ForeColor [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let Title(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarTitle = vData
50020     If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
50030         SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
50040     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Title [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get Title() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Title = ti.lpStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Title [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let PopupOnDemand(ByVal vData As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarPopupOnDemand = vData
50020     'If m_lTTHwnd <> 0 Then
50030     'End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "PopupOnDemand [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get PopupOnDemand() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     PopupOnDemand = mvarPopupOnDemand
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "PopupOnDemand [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let BackColor(ByVal vData As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarBackColor = vData
50020     If m_lTTHwnd <> 0 Then
50030         SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
50040     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "BackColor [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get BackColor() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     BackColor = mvarBackColor
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "BackColor [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let TipText(ByVal vData As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarTipText = vData
50020     ti.lpStr = vData
50030     If m_lTTHwnd <> 0 Then
50040         SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
50050     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "TipText [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get TipText() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     TipText = mvarTipText
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "TipText [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let Style(ByVal vData As ttStyleEnum)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarStyle = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Style [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get Style() As ttStyleEnum
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Style = mvarStyle
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Style [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
'//////////////////////////////////////////////////////
Public Property Let Centered(ByVal vData As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     mvarCentered = vData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Centered [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
Public Property Get Centered() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Centered = mvarCentered
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Centered [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

'X and Y are in Pixel so dont send vbTwips value
Public Sub Show(hwnd As Long, Optional x As Long = 0, Optional Y As Long = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     Dim pt As POINTAPI
50030     Dim ptTip As Long
50040     Dim ret As Long
50050
50060     With pt
50070         .x = x
50080         .Y = Y
50090     End With
50100
50110
50120
50130 '    ''now set our tooltip info structure
50140 '    With ti
50150 '        '//dont incude TTF_SUBCLASS for on demand
50160 '        ''if we want it centered, then set that flag
50170 '        If mvarCentered Then
50180 '            If mvarPopupOnDemand = False Then
50190 '                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
50200 '            Else
50210 '                .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP
50220 '            End If
50230 '        Else
50240 '            If mvarPopupOnDemand = False Then
50250 '                .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
50260 '            Else
50270 '                .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_TRANSPARENT
50280 '            End If
50290 '        End If
50300 '        'set the hwnd prop to our parent control's hwnd
50310 '        .hwnd = m_lParentHwnd
50320 '        .lId = m_lParentHwnd                               '0
50330 '        .hInstance = App.hInstance
50340 '        .lSize = Len(ti)
50350 '    End With
50360 '
50370 '    ''add the tooltip structure
50380 '    SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti
50390
50400
50410     ret = ClientToScreen(hwnd, pt)
50420
50430     ptTip = pt.Y * &H10000
50440     ptTip = ptTip + pt.x
50450
50460     ' These two messages will set the position of the tooltip:
50470     ret = SendMessage(m_lTTHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
50480     ret = SendMessage(m_lTTHwnd, TTM_TRACKACTIVATE, True, ti)
50490
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Show")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
Public Function CreateToolTip(ByVal ParentHwnd As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim lWinStyle As Long
50020     If m_lTTHwnd <> 0 Then
50030         DestroyWindow m_lTTHwnd
50040     End If
50050     m_lParentHwnd = ParentHwnd
50060
50070     ''create baloon style if desired
50080     If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
50090
50100     m_lTTHwnd = CreateWindowEx(0&, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            lWinStyle, _
            0&, _
            0&, _
            0&, _
            0&, _
            m_lParentHwnd, _
            0&, _
            0&, _
            0&)
50220
50230     ''now set our tooltip info structure
50240     With ti
50250         '//dont incude TTF_SUBCLASS for on demand
50260         ''if we want it centered, then set that flag
50270         If mvarCentered Then
50280             If mvarPopupOnDemand = False Then
50290                 .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
50300             Else
50310                 .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP
50320             End If
50330         Else
50340             If mvarPopupOnDemand = False Then
50350                 .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
50360             Else
50370                 .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_TRANSPARENT
50380             End If
50390         End If
50400
50410         'set the hwnd prop to our parent control's hwnd
50420         .hwnd = m_lParentHwnd
50430         .lId = m_lParentHwnd                               '0
50440         .hInstance = App.hInstance
50450         '.lpstr = ALREADY SET
50460         '.lpRect = lpRect
50470         .lSize = Len(ti)
50480     End With
50490
50500     ''add the tooltip structure
50510     SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti
50520
50530     ''if we want a title or we want an icon
50540     If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
50550         SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
50560     End If
50570
50580     If mvarForeColor <> Empty Then
50590         SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
50600     End If
50610
50620     If mvarBackColor <> Empty Then
50630         SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
50640     End If
50650
50660     SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime
50670     SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mvarDelayTime
50680
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "CreateToolTip")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub Destroy()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     If m_lTTHwnd <> 0 Then
50020         DestroyWindow m_lTTHwnd
50030     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsBalloonToolTip", "Destroy")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


