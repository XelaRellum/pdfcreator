VERSION 5.00
Begin VB.UserControl isExplorerBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   15360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12585
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   MouseIcon       =   "isExplorerBar.ctx":0000
   ScaleHeight     =   1024
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   839
   ToolboxBitmap   =   "isExplorerBar.ctx":0152
   Begin VB.VScrollBar m_ScrollBar 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox m_pChild 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   -4500
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   -4500
      Top             =   4560
   End
   Begin VB.Image imgbuttons 
      Height          =   510
      Left            =   -1500
      Picture         =   "isExplorerBar.ctx":0464
      Top             =   6240
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "isExplorerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************
'
' Control Name: isExplorerBar
'
' Author:       Fred.cpp
'               fred_cpp@msn.com
'
' Page:         http://mx.geocities.com/fred_cpp/isexplorerbar.htm
'
' Current
' Version:      1.91
'
' Description:  The Explorer on Windows XP has a beautiful
'               and nice looking bar in the left side of the
'               windows. It gives to Windows Explorer a very
'               intuitive usage and provides information
'               about the content of the current windows and
'               elements on it. I decided to make one similar
'               , where I can customize links, actions and
'               Info About the content of my own programs.
'               but It was a very difficult task a months ago
'               , so I gave up.
'               A month ago I saw the VBAccelerator ExplorerBar
'               and I loved It. nice visual Effects, good
'               functionality and a brunch of options and
'               customizations. I wanted to use It on my own
'               projects, but.... It has alot of dependences
'               I mean, a lot! more than 6 ocx or dll's were
'               required to run It in a New PC.
'               Then, I wanted to make my own Explorer Bar.
'               but I was going to include less options and
'               less visual effects. I only wanted to have a
'               control that would mimic some appareance
'               effects, like the background, links, and group
'               headers. so I started with this. in a few days
'               It growed up too much, from less than 120 code
'               lines to more than 1000, (that I think is to
'               much for a "small" project). Now this has more
'               than 2500 lines of code. I hope you enjoy.
'               comments, suggestions and of course, votes,
'               are wellcome and very apreciated.
'
' Features:     Single File Control.
'               Uses the REAL THEME Style (even If you Change
'               the XP Theme to something diferent than XP Luna
'               ( Example: Mac OS Themes or GNome themes See
'               Screenshots)
'               As Requested Olive and Metallic Schemes emulated.
'               If you have an OS that don't support themes
'               (Win9x , WinMe, Win2000, classic Style will be
'               used).
'               You Can add It to a project Easily.
'               No exotic objects and collections, the control
'               is controlled using a set of easy to understand
'               functions.
'               Lot of useful Events (for the most common mouse
'               actions.
'               I believe It Supports all the basic
'               functionality of the Explorer Bar.
'
' Notes:        I will brak some rules on the design of
'               usercontrols. I'll try to make this not so
'               heavy, no persistable data will be added, and
'               everything will be added at Runtime. Also I'll
'               try to make it a Single Control File.
'
' Work In Progress: I'm planning to add the Capability of add
'               more than a single Special Group (I don't like
'               that, but It's a suggestion :/
'
'
' Requeriments: Uses a ImageList Object to get the icons for the
'               small items (common controls 5 or 6, I've tested
'               with them both and works fine). so you need to
'               add a reference to COMCTL32.ocx in your project.
'               To enable the use of windows themes, you can
'               insert in your project call to InitCommonControls
'               or use the module ModMain.Bas Included in this
'               project (Borrowed from VBAccelerator). Also you
'               need include a manifest file (also you can use
'               the file I included in this project, and rename
'               it to be the same name of our exe. other solution
'               is to include it in a resource file. If you can
'               see the XP Visual Style in your other controls,
'               this control will also be drawn with the theme
'               style.
'
' Known Bugs:   In Some visual Styles, some ExplorerBar Parts
'               are not Defined. I'm trying to make a good
'               aproach for those cases, maybe use the window
'               frame part. with a close = expand button :/
'               I've seen some themes with this replacement,
'               But I It's still a wish.
'               When you set the theme style from Windows Classic
'               To any other theme, the control can't redraw
'               properly. I still don't know why. The api call's
'               don't report a failure, but the control Is not
'               updated. As soon As I found a way to fix It, I'll
'               post It.
'
' More Bugs:    If you found a Bug, please e-mail me.
'
' Updates:      Thanks for the Huge support for this control.
'               thanks to The people has sugested enhacements,
'               reported Bugs, and helped me to improve the
'               performance of this control.
'
' 2004:05:17    I got too much bugs for use the GetMessage API
'               So I stop Using It. Mouse wheel now Is not
'               supported. Also the control doesn't respond to the
'               theme change message. so won't update afther a
'               theme change. But the Tab Order Bug Is Fixed.
'               Also the nonactivate form bug.
'
' 2004:05:11    Olive and Silver Scheme Colors EMULATED; Explorer
'               Doesn't use the real theme data for drawing that
'               Shceme colors, so I also used a image.
'               the color's are not the real colors, but It's a
'               nice approach.
'
' 2004:05:09    The control updates his appareance when User
'               changes the windows theme (hehehe, without
'               subclassing).
'               Now ScrollBar Works Better and hand cursor works
'               afther a click(thanks Charles P.V.)
'               DetailsImage Added. Need some optimization, But
'               works fine.
'
' 2004:05:24    New Default Align is Left. Default font color for
'               Items is Buttontext.
'               Small nonredraw bug fixed.
'               supports Child Controls!
'
' 2004:05:25    Non redrawing afther change from nonthemed to
'               themed is fixed. maybe just need to add animation/
'               fade efects to finish up the control.
'
' 2004:05:26    Replaced imagelist.picture with extracticon.
'
' 2004:05:27    Start Using self subclassing by Paul Caton.
'               Implementing ScrollBars Enhacements and Theme
'               Change detect.
'               Mouse wheel Movement is Back
'
' 2004:06:21    I've Added the Font Property. And Now The text
'               can Include Far asian Languages (Need Feedback and
'               / or Help)
'               Clear Structure Can be called from a Itemclick
'               (thaks to Ademir Mazer Jr).
'
' 2004:06:23    Some Optimizations thanks to Roger Gilchrist
'               <rojagilkrist@hotmail.com> for the help and Also a
'               to Ferd(z) for the help In the VB5 compatibility.
'
' 2004:06:25    SetItemText and SetItemIcon bugfixes,
'               SetGroupCaption Function Added. Thanks to
'               Joerg Hohaus and Bios for the bug Reports
'
' 2004:07:07    New Functions and Improvements made by Joerg
'               Hohaus. This Update Is Fully done by him!
'
' 2004:07:09    Added support for VBAccelerator Imagelist by
'               Joerg Hohaus. Again, he has done all the code!
'
' 2004:08:12    Small fixes for bugs in emulated Olive and
'               Silver apareances.
'
' The code starts Here!

Option Explicit

'*************************************************************
'
'   Control Version:
'
Private Const strCurrentVersion = "1.91"
'**************************************


'*************************************************************
'
'   Private Constants
'
'**************************************
'Auxiliar Constants
Private Const RDW_INVALIDATE = &H1
Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&
Private Const S_OK = 0
Private Const HWND_DESKTOP = 0
Private Const AC_SRC_OVER = &H0
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

'Message ID's
Private Const WM_USER = &H400
Private Const WM_THEMECHANGED       As Long = &H31A
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_CTLCOLORSCROLLBAR  As Long = &H137
Private Const WM_VSCROLL            As Long = &H115
Private Const WM_HSCROLL            As Long = &H114
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOUSEHOVER         As Long = &H2A1
Private Const WM_SYSCOLORCHANGE     As Long = &H15 '21

'Tooltips Constants
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
Private Const TOOLTIPS_CLASSA = "tooltips_class32"
'Gradient Constants
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V  As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2
Private Const GRADIENT_FILL_OP_FLAG As Long = &HFF
'System Colors
Private Const COLOR_3DDKSHADOW As Long = 21
Private Const COLOR_BTNFACE As Long = 15
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_3DLIGHT As Long = 22
Private Const COLOR_BTNSHADOW As Long = 16
Private Const COLOR_ACTIVEBORDER As Long = 10
Private Const COLOR_ACTIVECAPTION As Long = 2
Private Const COLOR_APPWORKSPACE As Long = 12
Private Const COLOR_BACKGROUND As Long = 1
Private Const COLOR_BTNTEXT As Long = 18
Private Const COLOR_CAPTIONTEXT As Long = 9
Private Const COLOR_GRADIENTACTIVECAPTION As Long = 27
Private Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
Private Const COLOR_GRAYTEXT As Long = 17
Private Const COLOR_HIGHLIGHT As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const COLOR_HOTLIGHT As Long = 26
Private Const COLOR_INACTIVEBORDER As Long = 11
Private Const COLOR_INACTIVECAPTION As Long = 3
Private Const COLOR_INACTIVECAPTIONTEXT As Long = 19
Private Const COLOR_MENU As Long = 4
Private Const COLOR_MENUTEXT As Long = 7
Private Const COLOR_SCROLLBAR As Long = 0
Private Const COLOR_WINDOW As Long = 5
Private Const COLOR_WINDOWFRAME As Long = 6
Private Const COLOR_WINDOWTEXT As Long = 8
Private Const COLOR_3DFACE As Long = COLOR_BTNFACE
Private Const COLOR_3DHIGHLIGHT As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_DESKTOP As Long = COLOR_BACKGROUND
Private Const COLOR_BTNHILIGHT As Long = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DSHADOW As Long = COLOR_BTNSHADOW
Private Const COLOR_3DHILIGHT As Long = COLOR_BTNHIGHLIGHT
'Subclassing Constants
Private Const GWL_WNDPROC          As Long = -4
Private Const PATCH_05             As Long = 93                               'Table B (before) entry count
Private Const PATCH_09             As Long = 137                              'Table A (after) entry count

'*************************************************************
'
'   Required Type Definitions
'
'*************************************************************

Private Type POINT
   x As Long
   Y As Long
End Type

Private Type Size
   cx As Long
   cy As Long
End Type

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type RGB            'Required for color trnsform using RGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type MSG             'Windows Message Structure
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINT
End Type

Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type TRIVERTEX          'For gradient Drawing
    x As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Private Type DRAWTEXTPARAMS 'Required for DrawText
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Private Type BLENDFUNCTION  'Required for Alphablend API
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

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

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type TOOLINFO       'Tooltip Window Types
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As Rect
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Private Type SCROLLINFO     ' Scroll bar
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Type BarItem        'Item Structure
    sParent As String       'Parent group key
    key As String           'Key for external Access
    Index As Integer        'Index mainly for internal control
    Caption As String       '...you know
    Icon As Integer         'icon number.
    mRect As Rect           'rect in the control
    bOver As Boolean        'is the mouse over this?
    'iState As Integer       'Current State of the item
End Type

Private Type BarGroup
    Index As Integer        'Index for Internal Control'also external acces can be done with this, but It's easier to acces using the key than the index.
    key As String           'Key for external access
    Type As Integer         'Experimental. I'll try to set each group as normal, details, Special or with child controls
    Caption As String       'Need more Information?:/
    Icon As Picture         'Group Icon
    items() As BarItem      'Array Of Items
    iItemsCount As Integer  'Count of Items in the group
    bExpanded As Boolean    'Is the group Expanded?
    mRect As Rect           'Rect of the group header
    bOver As Boolean        'Control variable, is the mouse over this?
    iState As Integer       'Current State of the group
    lItemsHeight As Long    'Group items frame height
    pChild As PictureBox    'Picture that act's as child for the group (Experimental)
End Type

Private Type UxTheme        'Imported from a Cls File from VBAccelerator.com
    sClass As String        'And edited to keep the control in a single file.
    Part As Long            'I didn't used all the constant definitions where used
    State As Long           'in the original file, cuz I don't need them all
    hdc As Long             'But I added some others I need, like text offset
    hwnd As Long            'properties and UseTheme, to Detect If the draw was
    Left As Long            'succesfull or not, and then use classic windows Style
    Top As Long             'Drawing.
    Width As Long           'All the credits about the usage of UxTheme.dll defined on
    Height As Long          'cUxTheme.cls go for Steve at www.vbaccelerator.com
    Text As String
    TextAlign As DrawTextFlags
    IconIndex As Long
    hIml As Long
    RaiseError As Boolean
    UseThemeSize As Boolean
    UseTheme As Boolean
    TextOffset As Long
    RightTextOffset  As Long
End Type



'*************************************************************
'
'   Required Enums
'
'*************************************************************
Private Enum DrawTextAdditionalFlags
   DTT_GRAYED = &H1           '// draw a grayed-out string
End Enum

Private Enum THEMESIZE
    TS_MIN             '// minimum size
    TS_TRUE            '// size without stretching
    TS_DRAW            '// size that theme mgr will use to draw part
End Enum

Private Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Private Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

Enum GRADIENT_FILL_RECT
    FillHor = GRADIENT_FILL_RECT_H
    FillVer = GRADIENT_FILL_RECT_V
End Enum

Enum GRADIENT_TO_CORNER
    All
    TopLeft
    TopRight
    BottomLeft
    BottomRight
End Enum

Enum CRADIENT_DIRECTION
    DirectionSlash
    DirectionBackSlash
End Enum

Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

Private Enum DrawEdgeEdgeTypes
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENOUTER = &H2
    BDR_RAISEDINNER = &H4
    BDR_SUNKENINNER = &H8

    BDR_OUTER = (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = (BDR_RAISEDINNER Or BDR_SUNKENINNER)
    BDR_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    BDR_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)


    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
End Enum

Private Enum DrawEdgeBorderFlags
    BF_LEFT = &H1
    BF_TOP = &H2
    BF_RIGHT = &H4
    BF_BOTTOM = &H8

    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

    BF_DIAGONAL = &H10
    
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

    BF_MIDDLE = &H800         '/* Fill in the middle */
    BF_SOFT = &H1000          '/* For softer buttons */
    BF_ADJUST = &H2000        '/* Calculate the space left over */
    BF_FLAT = &H4000          '/* For flat rather than 3D borders */
    BF_MONO = &H8000          '/* For monochrome borders */
End Enum

'Message before, after or both
Private Enum eMsgWhen
  MSG_AFTER = 1
  MSG_BEFORE = 2
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

#If False Then
Private MSG_AFTER, MSG_BEFORE, MSG_BEFORE_AND_AFTER
#End If

'*************************************************************
'
'   API Call Declares
'
'*************************************************************

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As Rect, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal blendFunc As Long) As Boolean
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal hdc As Long, prc As Rect) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As Rect, pContentRect As Rect) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlag As Long, ByVal dwTextFlags2 As Long, pRect As Rect) As Long
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, ByVal hIml As Long, ByVal iImageIndex As Long) As Long
Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, prc As Rect, ByVal eSize As THEMESIZE, psz As Size) As Long
Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As DrawTextFlags, pBoundingRect As Rect, pExtentRect As Rect) As Long
Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pDestRect As Rect, ByVal uEdge As DrawEdgeEdgeTypes, ByVal uFlags As DrawEdgeBorderFlags, pContentRect As Rect) As Long
Private Declare Function IsThemePartDefined Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long) As Long
Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, prcImage As Rect) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long


'*************************************************************
'
'   Private Vars
'
'*************************************************************
'Subclassing Variables
Private nMsgCntB                   As Long                                    'Before msg table entry count
Private nMsgCntA                   As Long                                    'After msg table entry count
Private aMsgTblB()                 As Long                                    'Before msg table array
Private aMsgTblA()                 As Long                                    'After msg table array
Private nAddrSubclass              As Long                                    'The address of our WndProc
Private nAddrOriginal              As Long                                    'The address of the existing WndProc
Private sCode                      As String                                  'Binary subclass handler code string
'Control Variables
Private m_iTopOffset As Integer
Private m_cUxTheme As UxTheme
Private cGroups() As BarGroup
Private iGroups As Integer
Private m_objImageList As Object
Private iImgLType As Integer        'holds type of Imagelist
Private m_bOver As Boolean
Private m_NotOnUse As Long
Private m_GroupTextColor As OLE_COLOR
Private m_ItemTextColor As OLE_COLOR
Private m_GroupTextHoverColor As OLE_COLOR
Private m_ItemTextHoverColor As OLE_COLOR
Private m_GroupHoverColor As OLE_COLOR
Private m_bSpecialGroup As Boolean
Private m_SpecialGroup As BarGroup
Private m_SpecialGroupIcon As Picture
Private m_SpecialGroupBackground As Picture
Private m_bDetailsGroup As Boolean
Private m_DetailsGroup As BarGroup
Private m_DetailsGroupTittle As String
Private m_DetailsGroupText As String
Private m_DetailsRect As Rect
Private m_LastTextHeight As Long
Private m_Width As Long
'Private m_tempImg As PictureBox
Private m_AllowRedraw As Boolean
'Private WithEvents m_ScrollBar As VScrollBar
Private m_ttBackColor As Long 'properties for tooltip
Private m_ttTitle As String
Private m_ttForeColor As Long
Private m_ttParentControl As Object
Private m_ttIcon As ttIconType
Private m_ttCentered As Boolean
Private m_ttStyle As ttStyleEnum
Private m_ttlHwnd As Long
Private m_tti As TOOLINFO
Private bTrackMessages As Boolean
Private m_RedrawRect As Rect
Private m_DetailsPicture As StdPicture
Private m_ParentForm As Form
Private sThemeFile As String
Private sColorName As String
Private UxThemeText As Boolean
Private bEnableVBAcIml As Boolean

Private m_SelectedGroup As Long
Private m_SelectedItem As Long

'*************************************************************
'
'   Events Declares
'
'**************************************

Event MouseOver()
Event MouseOut()
Event GroupHover(sGroup As String)
Event GroupOut(sGroup As String)
Event ItemClick(sGroup As String, sItemKey As String)
Event GroupClick(ByVal Group As Long, bExpanded As Boolean)
Event ItemHover(sGroup As String, sItemKey As String)
Event ItemOut(sGroup As String, sItemKey As String)

'*************************************************************
'
' Paul Caton Subclassing system.
'   a Huge work I have to thank him for.
'
'*************************************************************
'
'Subclass handler - MUST be the first Public routine in this file.
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lngHwnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
'Parameters:
50020   'bBefore    Indicates whether the the message is being processed before or after the default handler - only really needed if a message is being subclassed before & after.
50030   'bHandled   Set this to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and optionaly, an 'after' callback
50040   'lReturn    Set this as per your intentions and requirements, see the MSDN documentation for each individual message type
50050   'hWnd       The window handle, should be the hWnd of the the User Control
50060   'uMsg       The message number
50070   'wParam     Message related data
50080   'lParam     Message related data
'Notes:
50100   'If you really, really know what you're doing, it's possible to change the values
50110   'of the last four parameters in a 'before' callback so that different values get
50120   'passed to the default handler.. and optionaly, the 'after' callback
50130     Dim tmpval As Integer
50141     Select Case uMsg
              Case WM_CTLCOLORSCROLLBAR
50160             'Stop this message
50170             uMsg = 0
50180         Case WM_MOUSEWHEEL
50190             'Wheel movement.
50200             'Debug.Print "Mouseweel: wParam= " & Hex(wParam) & " - lParam = " & Hex(lParam)
50210             If m_ScrollBar.Visible Then
50220                 If wParam = &H780000 Then
50230                 'wparam contains the direction the wheel was moved.
50240                     tmpval = m_ScrollBar.Value - 32
50250                     m_ScrollBar.Value = IIf((tmpval < m_ScrollBar.Min), _
                                        m_ScrollBar.Min, tmpval)
50270                 ElseIf wParam = &HFF880000 Then
50280                     tmpval = m_ScrollBar.Value + 32
50290                     m_ScrollBar.Value = IIf((tmpval > m_ScrollBar.Max), _
                                        m_ScrollBar.Max, tmpval)
50310                 End If
50320             End If
50330         Case WM_MOUSELEAVE
50340             Debug.Print "API Mouse Leave"
50350         Case WM_MOUSEHOVER
50360             Debug.Print "API Mouse Hover"
50370         Case WM_MOUSEMOVE
50380             'Debug.Print "WM_MOUSEMOVE: ", wParam, lParam
50390         Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
50400             'Redraw!
50410             DoEvents
50420             UserControl_Paint
50430     End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zSubclass_Proc")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


'======================================================================================================
'User Control's Subclass code
Private Sub Subclass_AddMsg(ByVal uMsg As Long, ByVal When As eMsgWhen)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   If When And eMsgWhen.MSG_BEFORE Then                                        'If Before
50020     'Add the message, pass the before table and before table message count variables ByRef
50030     Call zAddMsg(uMsg, aMsgTblB, nMsgCntB, eMsgWhen.MSG_BEFORE)
50040   End If
50050
50060   If When And eMsgWhen.MSG_AFTER Then                                         'If After
50070     'Add the message, pass the after table and after table message count variables ByRef
50080     Call zAddMsg(uMsg, aMsgTblA, nMsgCntA, eMsgWhen.MSG_AFTER)
50090   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Subclass_AddMsg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Delete the message from the msg table
Private Sub Subclass_DelMsg(ByVal uMsg As Long, ByVal When As eMsgWhen)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   If When And eMsgWhen.MSG_BEFORE Then                                        'If before
50020     'Delete the message, pass the Before table and before message count variables ByRef
50030     Call zDelMsg(uMsg, aMsgTblB, nMsgCntB, eMsgWhen.MSG_BEFORE)
50040   End If
50050
50060   If When And eMsgWhen.MSG_AFTER Then                                         'If After
50070     'Delete the message, pass the After table and after message count variables ByRef
50080     Call zDelMsg(uMsg, aMsgTblA, nMsgCntA, eMsgWhen.MSG_AFTER)
50090   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Subclass_DelMsg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Return whether we're running in the IDE. Public for general utility purposes
Private Function Subclass_InIDE() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Debug.Assert zSetTrue(Subclass_InIDE)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Subclass_InIDE")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'Start the subclassing
Private Function Subclass_Start() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Const PATCH_01 As Long = 18                                                 'Code buffer offset to the location of the relative address to EbMode
50020   Const PATCH_02 As Long = 68                                                 'Address of the previous WndProc
50030   Const PATCH_03 As Long = 78                                                 'Relative address of SetWindowsLong
50040   Const PATCH_06 As Long = 116                                                'Address of the previous WndProc
50050   Const PATCH_07 As Long = 121                                                'Relative address of CallWindowProc
50060   Const PATCH_0A As Long = 186                                                'Address of the owner object
50070   Const FUNC_EBM As String = "EbMode"                                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
50080   Const FUNC_SWL As String = "SetWindowLongA"                                 'SetWindowLong allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
50090   Const FUNC_CWP As String = "CallWindowProcA"                                'We use CallWindowProc to call the original WndProc
50100   Const MOD_VBA5 As String = "vba5"                                           'Location of the EbMode function if running VB5
50110   Const MOD_VBA6 As String = "vba6"                                           'Location of the EbMode function if running VB6
50120   Const MOD_USER As String = "user32"                                         'Location of the SetWindowLong & CallWindowProc functions
50130   Dim i          As Long                                                      'Loop index
50140   Dim s          As String
50150   Dim sHex       As String                                                    'Hex code string
50160
50170   'Protect against double calling of Subclass_Start without having performed a Subclass_Stop first
50180   Debug.Assert (nAddrSubclass = 0)
50190
50200   'Store the hex pair machine code representation in sHex
50210   sHex = "5589E583C4F85731C08945FC8945F8EB0EE8xxxxx01x83F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168xxxxx02x6AFCFF7508E8xxxxx03xEBE031D24ABFxxxxx04xB9xxxxx05xE82D000000C3FF7514FF7510FF750CFF750868xxxxx06xE8xxxxx07x8945FCC331D2BFxxxxx08xB9xxxxx09xE801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B8xxxxx0Ax508B00FF90A4070000C3"
50220
50230   'Convert the string from hex pairs to bytes and store in the ASCII string opcode buffer
50240   For i = 1 To Len(sHex) Step 2                                               'For each pair of hex characters
50250     sCode = sCode & ChrB$(Val("&H" & Mid$(sHex, i, 2)))                       'Convert a pair of hex characters to a byte and append to the ASCII string
50260   Next i                                                                      'Next pair
50270
50280   nAddrSubclass = StrPtr(sCode)                                               'Remember the address of the string code
50290
50300   If Subclass_InIDE Then
50310     Call CopyMemory(ByVal nAddrSubclass + 15, &H9090, 2)                      'Patch the jmp (EB0E) with two nop's (90) enabling the IDE breakpoint/stop checking code
50320
50330     i = zAddrFunc(MOD_VBA6, FUNC_EBM)                                         'Get the address of EbMode in vba6.dll
50340     If i = 0 Then                                                             'Found?
50350       i = zAddrFunc(MOD_VBA5, FUNC_EBM)                                       'VB5 perhaps, try vba5.dll
50360     End If
50370
50380     Debug.Assert i                                                            'Ensure the EbMode function was found
50390     Call zPatchRel(PATCH_01, i)                                               'Patch the relative address to the EbMode api function
50400   End If
50410
50420   nAddrOriginal = GetWindowLong(UserControl.hwnd, GWL_WNDPROC)                'Get the original window proc
50430   Call zPatchVal(PATCH_02, nAddrOriginal)                                     'Original WndProc address for CallWindowProc, call the original WndProc
50440   Call zPatchRel(PATCH_03, zAddrFunc(MOD_USER, FUNC_SWL))                     'Address of the SetWindowLong api function
50450   Call zPatchVal(PATCH_05, 0)                                                 'Initial before table entry count
50460   Call zPatchVal(PATCH_06, nAddrOriginal)                                     'Original WndProc address for SetWindowLong, unsubclass on IDE stop
50470   Call zPatchRel(PATCH_07, zAddrFunc(MOD_USER, FUNC_CWP))                     'Address of the CallWindowProc api function
50480   Call zPatchVal(PATCH_09, 0)                                                 'Initial after table entry count
50490   Call zPatchVal(PATCH_0A, ObjPtr(Me))                                        'Get the address of the current instance of this User Control
50500   nAddrOriginal = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, nAddrSubclass) 'Set our WndProc in place of the original
50510
50520   If nAddrOriginal <> 0 Then
50530     Subclass_Start = True                                                     'Success
50540   End If
50550
50560   Debug.Assert Subclass_Start
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Subclass_Start")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'Stop subclassing
Private Sub Subclass_Stop()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Debug.Assert nAddrSubclass                                                  'Ensure that we are subclassing before we attempt to stop
50020   Call zPatchVal(PATCH_05, 0)                                                 'Patch the Table B entry count to ensure no further 'before' callbacks
50030   Call zPatchVal(PATCH_09, 0)                                                 'Patch the Table A entry count to ensure no further 'after' callbacks
50040   Call SetWindowLong(UserControl.hwnd, GWL_WNDPROC, nAddrOriginal)            'Restore the original WndProc
50050   nMsgCntB = 0                                                                'Message before count set to zero
50060   nMsgCntA = 0                                                                'Message after count set to zero
50070   nAddrSubclass = 0                                                           'Indicate that we aren't subclassing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Subclass_Stop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'======================================================================================================
'These "z" routines are used by the subclass code - they shouldn't be called directly by the control author

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Const PATCH_04 As Long = 88                                                 'Table B (before) address
50020   Const PATCH_08 As Long = 132                                                'Table A (after) address
50030   Dim nEntry     As Long
50040   Dim nOff1      As Long
50050   Dim nOff2      As Long
50060
50070   If uMsg = -1 Then                                                           'If all messages
50080     nMsgCnt = -1                                                              'Indicates that all messages shall callback
50090   Else                                                                        'Else a specific message number
50100     For nEntry = 1 To nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
50111       Select Case aMsgTbl(nEntry)                                             'Select on the message number stored in this table entry
            Case -1                                                                 'This msg table slot is a deleted entry
50130         aMsgTbl(nEntry) = uMsg                                                'Re-use this entry
50140         Exit Sub                                                              'Bail
50150       Case uMsg                                                               'The msg is already in the table!
50160         Exit Sub                                                              'Bail
50170       End Select
50180     Next nEntry                                                               'Next entry
50190
50200     'Make space for the new entry
50210     ReDim Preserve aMsgTbl(1 To nEntry)                                       'Increase the size of the table. NB nEntry = nMsgCnt + 1
50220     nMsgCnt = nEntry                                                          'Bump the entry count
50230     aMsgTbl(nEntry) = uMsg                                                    'Store the message number in the table
50240   End If
50250
50260   If When = eMsgWhen.MSG_BEFORE Then                                          'If before
50270     nOff1 = PATCH_04                                                          'Offset to the Before table address
50280     nOff2 = PATCH_05                                                          'Offset to the Before table entry count
50290   Else                                                                        'Else after
50300     nOff1 = PATCH_08                                                          'Offset to the After table address
50310     nOff2 = PATCH_09                                                          'Offset to the After table entry count
50320   End If
50330
50340   'Patch the appropriate table entries
50350   Call zPatchVal(nOff1, zAddrMsgTbl(aMsgTbl))                                 'Patch the appropriate table address. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a message number is added.
50360   Call zPatchVal(nOff2, nMsgCnt)                                              'Patch the appropriate table entry count
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zAddMsg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Return the address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   zAddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
50020
50030   'You may want to comment out the following line if you're using vb5 else the EbMode
50040   'GetProcAddress will stop here everytime because we look in vba6.dll first
50050   Debug.Assert zAddrFunc
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zAddrFunc")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'Return the address of the low bound of the passed table array
Private Function zAddrMsgTbl(ByRef aMsgTbl() As Long) As Long
  On Error Resume Next                                                        'The table may not be dimensioned yet so we need protection
  zAddrMsgTbl = VarPtr(aMsgTbl(1))                                            'Get the address of the first element of the passed message table
  On Error GoTo 0                                                             'Switch off error protection
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim nEntry As Long
50020
50030   If uMsg = -1 Then                                                           'If deleting all messages
50040     nMsgCnt = 0                                                               'Message count is now zero
50050     If When = eMsgWhen.MSG_BEFORE Then                                        'If before
50060       nEntry = PATCH_05                                                       'Patch the before table message count location
50070     Else                                                                      'Else after
50080       nEntry = PATCH_09                                                       'Patch the after table message count location
50090     End If
50100     Call zPatchVal(nEntry, 0)                                                 'Patch the table message count
50110   Else                                                                        'Else deleteting a specific message
50120     For nEntry = 1 To nMsgCnt                                                 'For each table entry
50130       If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
50140         aMsgTbl(nEntry) = -1                                                  'Mark the table slot as available
50150         Exit For                                                              'Bail
50160       End If
50170     Next nEntry                                                               'Next entry
50180   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zDelMsg")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Patch the machine code buffer offset with the relative address to the target address
Private Sub zPatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Call CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zPatchRel")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Patch the machine code buffer offset with the passed value
Private Sub zPatchVal(ByVal nOffset As Long, ByVal nValue As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Call CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zPatchVal")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Worker function for Subclass_InIDE - will only be called whilst running in the IDE
Private Function zSetTrue(bValue As Boolean) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   zSetTrue = True
50020   bValue = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "zSetTrue")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function




'*************************************************************
'
'   User control Events
'
'*************************************************************

' Desc: Read the properties from the property bag -
'       also, a good place to start the subclassing
'       (if we're running) - this could also be enabled for
'       design time... if that's what you want.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With m_cUxTheme
        .hdc = UserControl.hdc
        .Width = 120
        .Height = 24
        .TextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
    End With
    m_ItemTextColor = 0
    m_GroupTextColor = 0
    m_NotOnUse = 0
    m_ItemTextHoverColor = RGB(127, 127, 127)
    m_GroupHoverColor = RGB(127, 127, 127)
    UserControl.Extender.Align = 3
    m_ttIcon = TTIconInfo
    m_ttTitle = App.Title
    With PropBag
        UserControl.Fontname = .ReadProperty("FontName", UserControl.Ambient.Font.Name)
        UserControl.Font.Charset = .ReadProperty("FontCharset")
        UxThemeText = CBool(.ReadProperty("UxThemeText", True))
        bEnableVBAcIml = CBool(.ReadProperty("EnableVBAcIml", False))
    End With
  
    'If we're not in design mode
    If Ambient.UserMode Then
        'Start subclassing
        Call Subclass_Start

        'Add the messages that we're interested in
        Call Subclass_AddMsg(WM_THEMECHANGED, MSG_AFTER)
        Call Subclass_AddMsg(WM_SYSCOLORCHANGE, MSG_AFTER)
        Call Subclass_AddMsg(WM_MOUSEMOVE, MSG_AFTER)
        Call Subclass_AddMsg(WM_CTLCOLORSCROLLBAR, MSG_BEFORE)
        Call Subclass_AddMsg(WM_MOUSEWHEEL, MSG_AFTER)
        Call Subclass_AddMsg(WM_MOUSELEAVE, MSG_AFTER)
        Call Subclass_AddMsg(WM_MOUSEHOVER, MSG_AFTER)
    End If
End Sub

' Desc: Save the properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     With PropBag
50020         .WriteProperty "FontName", UserControl.Font.Name, "Verdana"
50030         .WriteProperty "FontCharset", UserControl.Font.Charset
50040         .WriteProperty "UxThemeText", UxThemeText, True
50050     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_WriteProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Initialize control
Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     bEnableVBAcIml = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   If nAddrSubclass <> 0 Then                                                  'If we're subclassing
50020     Call Subclass_Stop                                                        'Stop subclassing
50030   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_Terminate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: when the scrollbar is visible and changes,
'   update the offset and redraw contents
Private Sub m_ScrollBar_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_iTopOffset = m_ScrollBar.Value
50020     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "m_ScrollBar_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub m_pChild_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UserControl_MouseMove 0, Shift, 3, 3
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "m_pChild_MouseMove")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_AmbientChanged")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: when the scrollbar is visible and changes,
'   update the offset and redraw contents
Private Sub m_ScrollBar_Scroll()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_ScrollBar_Change
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "m_ScrollBar_Scroll")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: to draw the apropiated background to the child
'       control, I'll try to caught that events here
Private Sub m_pChild_Paint(Index As Integer)
    'Child picturteboxes were redirected to this control
    'pChild(Index).hdc
    On Error Resume Next
    Dim nj As Integer
    With m_cUxTheme
    If cGroups(Index).bExpanded Then
        If Not cGroups(Index).pChild Is Nothing Then
            'Child Picture Box Is Defined!
            Dim ltmpBackColor As Long
            cGroups(Index).pChild.Move cGroups(Index).mRect.Left * Screen.TwipsPerPixelX, (cGroups(Index).mRect.Bottom) * Screen.TwipsPerPixelY, (cGroups(Index).mRect.Right - cGroups(Index).mRect.Left) * Screen.TwipsPerPixelX
            cGroups(Index).pChild.Visible = True
            cGroups(Index).pChild.AutoRedraw = True
            .hdc = cGroups(Index).pChild.hdc
            .hwnd = cGroups(Index).pChild.hwnd
            .Left = 0: .Top = 0: .Width = cGroups(Index).pChild.ScaleWidth: .Height = cGroups(Index).pChild.ScaleHeight
            .Part = 5
            .State = 1
            .Text = ""
            .Part = 5
            .State = 1
            Select Case sColorName
            'this styles now are EMULATED. (just like microsoft does)
                Case "Metallic"
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), vbWhite, B
                    ltmpBackColor = RGB(&HF0, &HF1, &HF5)
                Case "HomeStead"
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), vbWhite, B
                    ltmpBackColor = RGB(&HF6, &HF6, &HEC)
                Case "Classic"
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), GetSysColor(COLOR_WINDOW), BF
                    cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), GetSysColor(COLOR_BTNFACE), B
                    ltmpBackColor = GetSysColor(COLOR_WINDOW)
                Case Else
                DrawTheme
                    ltmpBackColor = GetPixel(cGroups(Index).pChild.hdc, 4, 4) ' RGB(&HF0, &HF1, &HF5)
            End Select
            If Not .UseTheme Then
                'Draw Failed, use Classic Style
                cGroups(Index).pChild.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbButtonFace, B
            End If
            Dim tmpCtl
            For Each tmpCtl In UserControl.ParentControls
                On Error Resume Next
                If tmpCtl.Container.Name = m_pChild(Index).Tag Then
                    If TypeOf tmpCtl Is OptionButton Then
                        'Is an option button?
                        tmpCtl.BackColor = ltmpBackColor
                    ElseIf TypeOf tmpCtl Is Label Then
                        'Is a Label?
                        tmpCtl.BackColor = ltmpBackColor
                    ElseIf TypeOf tmpCtl Is CheckBox Then
                        'Is a Checkbox?
                        tmpCtl.BackColor = ltmpBackColor
                    End If
                End If
            Next
            .hdc = UserControl.hdc
            .hwnd = UserControl.hwnd
        Else
            'hide the child picturebox
            'cGroups(Index).pChild.Move cGroups(Index).mRect.Left * Screen.TwipsPerPixelX, (cGroups(Index).mRect.Bottom) * Screen.TwipsPerPixelY, (cGroups(Index).mRect.Right - cGroups(Index).mRect.Left) * Screen.TwipsPerPixelX
            cGroups(Index).pChild.Visible = False
        'group has been drawn
        End If
    End If
    End With
End Sub

' desc: Here we process when the user Pushes
'       over items and header groups.
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Check for Click events
50020     'Process the current Events
50030     Dim ni As Integer, nj As Integer
50040     'Currently only left button actions supported.
50050     'If the button is different from vbleftbutton
50060     If Button = vbLeftButton Then
50070     'Check in the existing objects to see if anyone
50080     'has been presed
50090     If m_bSpecialGroup Then
50100         If Y >= m_SpecialGroup.mRect.Top And Y <= m_SpecialGroup.mRect.Bottom And m_SpecialGroup.mRect.Left < x And m_SpecialGroup.mRect.Right > x Then
50110             'Mouse Down! Redraw Group Header
50120             m_SpecialGroup.iState = 3
50130             RedrawSpecialHeader
50140         End If
50150         If m_SpecialGroup.bExpanded Then
50160             'Analice each item for the group
50170             For nj = 1 To m_SpecialGroup.iItemsCount
50180                 'Search each item
50190                 If Y >= m_SpecialGroup.items(nj).mRect.Top And Y <= m_SpecialGroup.items(nj).mRect.Bottom And m_SpecialGroup.items(nj).mRect.Left < x And m_SpecialGroup.items(nj).mRect.Right > x Then
50200                     'Item down
50210                     m_SelectedGroup = -1
50220                     m_SelectedItem = nj
50230                     RedrawItem -1, nj, 3
50240                 End If
50250             Next nj
50260         End If
50270     End If
50280     'Normal Groups
50290     For ni = 1 To iGroups
50300         If Y >= cGroups(ni).mRect.Top And Y <= cGroups(ni).mRect.Bottom And cGroups(ni).mRect.Left < x And cGroups(ni).mRect.Right > x Then
50310             'Mouse Down! Redraw Group Header
50320             cGroups(ni).iState = 3
50330             RedrawGroupHeader ni
50340         End If
50350         If cGroups(ni).bExpanded Then
50360             'Analice each item for the group
50370             For nj = 1 To cGroups(ni).iItemsCount
50380                 'Search each item
50390                 If Y >= cGroups(ni).items(nj).mRect.Top And Y <= cGroups(ni).items(nj).mRect.Bottom And cGroups(ni).items(nj).mRect.Left < x And cGroups(ni).items(nj).mRect.Right > x Then
50400                     'Item down
50410                     m_SelectedGroup = ni
50420                     m_SelectedItem = nj
50430                     RedrawItem ni, nj, 3
50440                 End If
50450             Next nj
50460         End If
50470     Next ni
50480     'Details Group
50490     If m_bDetailsGroup Then
50500         If Y >= m_DetailsGroup.mRect.Top And Y <= m_DetailsGroup.mRect.Bottom And m_DetailsGroup.mRect.Left < x And m_DetailsGroup.mRect.Right > x Then
50510             'Mouse Down! Redraw Group Header
50520             m_DetailsGroup.iState = 3
50530             RedrawDetailsHeader
50540         End If
50550     End If
50560     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_MouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: when the mouse pointer moves over the control,
'       some controls will be highlighted, other
'       deactivated. here we can process that events.
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Find out the Area where the mouse is located and highlight the current "object"
50020     Dim OldOver As Boolean
50030     Dim ni As Integer, nj As Integer
50040
50050     'OldOver = m_bOver   'Set Previous State
50060     'm_bOver = (x > 0) And (y > 0) And (x < UserControl.ScaleWidth) And (y < UserControl.ScaleHeight)
50070     'm_bOver = (X > 0) And (Y > 0) And (X < UserControl.ScaleWidth - IIf(m_ScrollBar.Visible, m_ScrollBar.Width, 0)) And (Y < UserControl.ScaleHeight)
50080     'If (m_bOver And Not OldOver) Then
50090     '    RaiseEvent MouseOver
50100     '    Debug.Print "Mouse Over!"
50110     'End If
50120     If UserControl.Enabled And Button = 0 Then
50130         If Not timUpdate.Enabled Then
50140             timUpdate.Enabled = True
50150             UserControl_MouseMove 0, 0, 1, 1
50160         Else
50170             timUpdate.Enabled = True
50180         End If
50190     End If
50200     DoEvents
50210     'Process the current Events
50220     'Start on the Special Group
50230     If m_bSpecialGroup Then
50240         If Y >= m_SpecialGroup.mRect.Top And Y <= m_SpecialGroup.mRect.Bottom And m_SpecialGroup.mRect.Left < x And m_SpecialGroup.mRect.Right > x Then
50250             'Cursor is over the group
50260             If Not m_SpecialGroup.bOver Then
50270                 m_SpecialGroup.bOver = True
50280                 m_SpecialGroup.iState = 2
50290                 SetHandCur True
50300                 RedrawSpecialHeader
50310                 'Raise Event for this group
50320                 RaiseEvent GroupHover(m_SpecialGroup.key)
50330                 'Debug.Print "Over Special Group "
50340                 'm_ttTitle = "Tittle"
50350                 'm_tti.lpStr = "Tooltip data"
50360             End If
50370         Else    'cursor is not over the group
50380             'Was In? then set out
50390             If m_SpecialGroup.bOver Then
50400                 m_SpecialGroup.bOver = False
50410                 m_SpecialGroup.iState = 1
50420                 SetHandCur False
50430                 RedrawSpecialHeader
50440                 'Raise Event for this group
50450                 RaiseEvent GroupOut(m_SpecialGroup.key)
50460                 'Debug.Print "Exit Special Group "
50470             End If
50480         End If
50490         If m_SpecialGroup.bExpanded Then
50500             'Analice each item for the group
50510             For nj = 1 To m_SpecialGroup.iItemsCount
50520                 'Search each item
50530                 If Y >= m_SpecialGroup.items(nj).mRect.Top And Y <= m_SpecialGroup.items(nj).mRect.Bottom And m_SpecialGroup.items(nj).mRect.Left < x And m_SpecialGroup.items(nj).mRect.Right > x Then
50540                     'Cursor Hover the item
50550                     If Not m_SpecialGroup.items(nj).bOver Then
50560                         'Set Hover
50570                         m_SpecialGroup.items(nj).bOver = True
50580                         RedrawItem -1, nj, 2
50590                         SetHandCur True
50600                         RaiseEvent ItemHover(m_SpecialGroup.key, m_SpecialGroup.items(nj).key)
50610                         'Debug.Print "Hover Item: " & nj
50620                     End If
50630                 Else
50640                     'Was Over this item?
50650                     If m_SpecialGroup.items(nj).bOver Then
50660                         'Set Out
50670                         m_SpecialGroup.items(nj).bOver = False
50680                         RedrawItem -1, nj, 1
50690                         SetHandCur False
50700                         RaiseEvent ItemOut(m_SpecialGroup.key, m_SpecialGroup.items(nj).key)
50710                         'Debug.Print "Out Special Item: " & nj
50720                     End If
50730                 End If
50740             Next nj
50750         End If
50760     End If
50770
50780     ''Search in the normal groups
50790     For ni = 1 To iGroups
50800         If Y >= cGroups(ni).mRect.Top And Y <= cGroups(ni).mRect.Bottom And cGroups(ni).mRect.Left < x And cGroups(ni).mRect.Right > x Then
50810             'Cursor is over the group
50820             If Not cGroups(ni).bOver Then
50830                 cGroups(ni).bOver = True
50840                 cGroups(ni).iState = 2
50850                 RedrawGroupHeader ni
50860                 'Raise Event for this group
50870                 SetHandCur True
50880                 RaiseEvent GroupHover(cGroups(ni).key)
50890                 'Debug.Print "over Group " & ni
50900             End If
50910         Else    'cursor is not over the group
50920             'Was In? then set out
50930             If cGroups(ni).bOver Then
50940                 cGroups(ni).bOver = False
50950                 cGroups(ni).iState = 1
50960                 RedrawGroupHeader ni
50970                 SetHandCur False
50980                 'Raise Event for this group
50990                 RaiseEvent GroupOut(cGroups(ni).key)
51000                 'Debug.Print "Exit Group " & ni
51010             End If
51020         End If
51030         If cGroups(ni).bExpanded Then
51040             'Analice each item for the group
51050             For nj = 1 To cGroups(ni).iItemsCount
51060                 'Search each item
51070                 If Y >= cGroups(ni).items(nj).mRect.Top And Y <= cGroups(ni).items(nj).mRect.Bottom And cGroups(ni).items(nj).mRect.Left < x And cGroups(ni).items(nj).mRect.Right > x Then
51080                     'Cursor Hover the item
51090                     If Not cGroups(ni).items(nj).bOver Then
51100                         'Set Hover
51110                         cGroups(ni).items(nj).bOver = True
51120                         RedrawItem ni, nj, 2
51130                         SetHandCur True
51140                         'Raiseevent ItemOver
51150                         RaiseEvent ItemHover(cGroups(ni).key, cGroups(ni).items(nj).key)
51160                         'Debug.Print "Hover Item: " & nj
51170                     End If
51180                 Else
51190                     'Was Over this item?
51200                     If cGroups(ni).items(nj).bOver Then
51210                         'Set Out
51220                         cGroups(ni).items(nj).bOver = False
51230                         RedrawItem ni, nj, 1
51240                         SetHandCur False
51250                         'Raiseevent ItemOut
51260                         RaiseEvent ItemOut(cGroups(ni).key, cGroups(ni).items(nj).key)
51270                         'Debug.Print "Out Item: " & nj
51280                     End If
51290                 End If
51300             Next nj
51310         End If
51320     Next ni
51330     'Search on the Details
51340     If m_bDetailsGroup Then
51350         If Y >= m_DetailsGroup.mRect.Top And Y <= m_DetailsGroup.mRect.Bottom And m_DetailsGroup.mRect.Left < x And m_DetailsGroup.mRect.Right > x Then
51360             'Cursor is over the group
51370             If Not m_DetailsGroup.bOver Then
51380                 m_DetailsGroup.bOver = True
51390                 m_DetailsGroup.iState = 2
51400                 SetHandCur True
51410                 RedrawDetailsHeader
51420                 'Raise Event for this group
51430                 RaiseEvent GroupHover(m_DetailsGroup.key)
51440                 'Debug.Print "Over Details Group "
51450             End If
51460         Else    'cursor is not over the group
51470             'Was In? then set out
51480             If m_DetailsGroup.bOver Then
51490                 m_DetailsGroup.bOver = False
51500                 m_DetailsGroup.iState = 1
51510                 SetHandCur False
51520                 RedrawDetailsHeader
51530                 'Raise Event for this group
51540                 RaiseEvent GroupOut(m_DetailsGroup.key)
51550                 'Debug.Print "Exit Details Group "
51560             End If
51570         End If
51580     End If
51590
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_MouseMove")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
' Desc: The clicks on the objects of the control,
'       are raised here, when the user releases
'       the button.
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Check for Click events
    'Process the current Events
    Dim ni As Integer, nj As Integer
    'Small Fix to allow clearStructure from an ItemClick event
    ' Thanks to Ademir Mazer Jr
    Dim GroupKeyAux As String
    Dim ItemKeyAux As String
    'Currently only left button actions supported.
    'If the button is different from vbleftbutton
    'then exit this sub.
    If Button = vbLeftButton Then
        On Error GoTo ItemDoesntExist
        'Search in special group
        If m_bSpecialGroup Then
            If Y >= m_SpecialGroup.mRect.Top And Y <= m_SpecialGroup.mRect.Bottom And m_SpecialGroup.mRect.Left < x And m_SpecialGroup.mRect.Right > x Then
                'Cursor is over the group
                m_SpecialGroup.bExpanded = Not m_SpecialGroup.bExpanded
                m_SpecialGroup.iState = 2
                UserControl_Paint
                'SetHandCur True
                RaiseEvent GroupClick(-1, m_SpecialGroup.bExpanded)
            End If
            'Analice each item for the group
            If m_SpecialGroup.bExpanded Then
                For nj = 1 To m_SpecialGroup.iItemsCount
                    'Search each item
                    If Y >= m_SpecialGroup.items(nj).mRect.Top And Y <= m_SpecialGroup.items(nj).mRect.Bottom And m_SpecialGroup.items(nj).mRect.Left < x And m_SpecialGroup.items(nj).mRect.Right > x Then
                        'Cursor Hover the item
                        RedrawItem -1, nj, 2
                        'Small Fix to allow clearStructure from an ItemClick event
                        ' Thanks to Ademir Mazer Jr
                        GroupKeyAux = m_SpecialGroup.key
                        ItemKeyAux = m_SpecialGroup.items(nj).key
                        RaiseEvent ItemClick(GroupKeyAux, ItemKeyAux)
                    End If
                Next nj
            End If
        End If
        
        'Search the normal groups
        For ni = 1 To iGroups
            If Y >= cGroups(ni).mRect.Top And Y <= cGroups(ni).mRect.Bottom And cGroups(ni).mRect.Left < x And cGroups(ni).mRect.Right > x Then
                'Cursor is over the group
                cGroups(ni).bExpanded = Not cGroups(ni).bExpanded
                cGroups(ni).iState = 2
                UserControl_Paint
                UserControl.Refresh
                'SetHandCur True
                RaiseEvent GroupClick(ni, cGroups(ni).bExpanded)
            End If
            'Analice each item for the group
            If cGroups(ni).bExpanded Then
                For nj = 1 To cGroups(ni).iItemsCount
                    'Search each item
                    If Y >= cGroups(ni).items(nj).mRect.Top And Y <= cGroups(ni).items(nj).mRect.Bottom And cGroups(ni).items(nj).mRect.Left < x And cGroups(ni).items(nj).mRect.Right > x Then
                        'Cursor Hover the item
                        RedrawItem ni, nj, 2
                        GroupKeyAux = cGroups(ni).key
                        ItemKeyAux = cGroups(ni).items(nj).key
                        RaiseEvent ItemClick(GroupKeyAux, ItemKeyAux)
                    End If
                Next nj
            End If
        Next ni
        
        'Search in Details group
        If m_bDetailsGroup Then
            If Y >= m_DetailsGroup.mRect.Top And Y <= m_DetailsGroup.mRect.Bottom And m_DetailsGroup.mRect.Left < x And m_DetailsGroup.mRect.Right > x Then
                'Cursor is over the group
                m_DetailsGroup.bExpanded = Not m_DetailsGroup.bExpanded
                m_DetailsGroup.iState = 2
                UserControl_Paint
                UserControl.Refresh
                'SetHandCur True
                RaiseEvent GroupClick(-2, m_DetailsGroup.bExpanded)
            End If
        End If
    End If
ItemDoesntExist:
    Call UserControl_MouseMove(Button, Shift, x, Y)
End Sub

' Desc: When the control is resized, redraw everything
Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: This sub Is executed when the control Is shown
'       I added code to detect some messages that VB
'       don't notify.
Private Sub UserControl_Show()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_AllowRedraw = True
50020     UserControl_Paint
50030
50040 '   please see http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=52168&lngWId=1
50050 '   Also thanks to Min Thant Sin at 7:33:33 AM on 5/3/2004
50060 '   see http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngTopicId=31065&lngWId=1&Forum=Visualbasic&TopicCategory=%20Request%20for%20Code
50070 '    'I used this to track some messages. But this feature generated
50080 '    '   too many bugs. So I quit. If you found this code usefull, you can use It.
50090 '    If UserControl.Ambient.UserMode Then
50100 '        bTrackMessages = True
50110 '        Do Until bTrackMessages = False
50120 '            DoEvents
50130 '            Call TrackMessage
50140 '            DoEvents
50150 '        Loop
50160 '    End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UserControl_Show")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Refresh")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
' Desc: This sub Is used to know If the mouse is
'       inside the control.
Private Sub timUpdate_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Check Out If the mouse is inside the control.
50020     If InBox(UserControl.hwnd) Then
50030         If m_bOver = False Then
50040             UserControl_Paint
50050             RaiseEvent MouseOver
50060         End If
50070         m_bOver = True
50080     Else
50090         If m_bOver Then
50100             'UserControl_Paint
50110             timUpdate.Enabled = False
50120             RaiseEvent MouseOut
50130         End If
50140         m_bOver = False
50150         'If any object was highlighted, reset all.
50160         UserControl_MouseMove 0, 0, 1, 1
50170     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "timUpdate_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


' Desc: This sub is where I draw the control objects.
'       Everything is here. maybe you can learn some
'       Things from here. I learned a lot from
'       VBThemeEplorer in vbaccelerator. the code for
'       drawing using UxTheme comes from that project.
'       but I turned the class into a structure and
'       changed his method Draw into a function (
'       Drawtheme), so, Now I don't need a extra file.
Private Sub UserControl_Paint()
    Dim ni As Integer, nj As Integer
    Dim iTop As Integer
    Dim bUseTheme As Boolean
    Dim tmpRect As Rect
    If Not m_AllowRedraw Then Exit Sub
    If Not UserControl.Ambient.UserMode Then
        'Stop filetring Messages
        bTrackMessages = False
        'Draw a Nice Banner :P
        With m_cUxTheme
            'Setup Some properties
            .hdc = UserControl.hdc
            .hwnd = UserControl.hwnd
            .sClass = "Explorerbar"
            .Part = 1
            .State = 1
            .Left = 0
            .Top = 0
            .Width = UserControl.Width
            .Height = UserControl.Height
            'Draw Background
            DrawTheme
            .Part = 9
            .Left = 3
            .Width = UserControl.ScaleWidth - 6
            .Height = 60
            .TextOffset = 0
            .RightTextOffset = 0
            .Top = 48
            .Text = "http://mx.geocities.com/fred_cpp/isexplorerbar"
            .TextAlign = DT_CENTER Or DT_TOP Or DT_WORD_ELLIPSIS
            DrawTheme
            .Part = 12
            .Top = 25
            '.Left = 30
            '.Width = UserControl.ScaleWidth - 60
            .Height = 24
            .State = 2
            .TextAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
            .Text = "isExplorerBar"
            DrawTheme
            'Dim tmpRect As RECT
            'SetRect tmpRect, 10, .Top + 4, UserControl.Width - 10, .Top + 48
            'DrawRectText tmpRect, "http://mx.geocities.com/fred_cpp/isexplorerbar"
            
            If Not .UseTheme Then
                'No theme aviable, use classic drawing
                UserControl.Cls
                SetRect tmpRect, 8, 12, UserControl.Width - 24, 34
                UserControl.Line (6, 12)-(UserControl.ScaleWidth - 12, 34), vbHighlight, BF
                UserControl.ForeColor = vbHighlightText
                UserControl.FontBold = True
                DrawText UserControl.hdc, "isExplorerBar", -1, tmpRect, DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_MODIFYSTRING
                SetRect tmpRect, 8, 35, UserControl.Width - 24, 88
                UserControl.Line (6, 35)-(UserControl.ScaleWidth - 12, 88), vbHighlight, B
                UserControl.ForeColor = vbButtonText
                UserControl.FontBold = False
                DrawText UserControl.hdc, "http://mx.geocities.com/fred_cpp/" & vbCrLf & "isexplorerbar", -1, tmpRect, DT_WORD_ELLIPSIS Or DT_MODIFYSTRING
            End If
            'PaintPicture toolboxbitmap?:( It's not possible? :/
        End With
    Else
        'Calculate the position and rects for each item.
        CalcRects
        'Get the theme name
        GetThemeName
        With m_cUxTheme
            'Setup Some properties
            .hdc = UserControl.hdc
            .hwnd = UserControl.hwnd
            .sClass = "Explorerbar"
            .Part = 1
            .State = 1
            .Left = 0
            .Top = 0
            .Width = UserControl.Width
            .Height = UserControl.Height
            .TextOffset = 32
            .RightTextOffset = 25
            .TextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
            Select Case sColorName
                '<wip>Background is still not exactly the same</wip>
                'Case "Metallic"
                    'DoGradient RGB(&HC3, &HC7, &HD3), RGB(&HB1, &HB3, &HC8), FillVer, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
                    'UserControl.BackColor = RGB(&HC3, &HC7, &HD3)
                    'UserControl.Cls
                    'DrawTheme
                Case "Classic"
                    'No theme aviable, use classic drawing
                    UserControl.Cls
                Case Else
                    'Other
                    DrawTheme
            End Select
            'Check for the Special Group
            If m_bSpecialGroup Then
                'Draw the Special group
                RedrawSpecialHeader
                If m_SpecialGroup.bExpanded Then
                    'Draw the group Items frame
                    .Part = 9
                    .State = 1
                    .Text = ""
                    .Left = m_SpecialGroup.mRect.Left
                    .Top = m_SpecialGroup.mRect.Bottom
                    .Height = m_SpecialGroup.lItemsHeight
                    .Width = m_SpecialGroup.mRect.Right - m_SpecialGroup.mRect.Left
                    Select Case sColorName
                    'this styles now are EMULATED. (just like microsoft does)
                        Case "Metallic"
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                        Case "HomeStead"
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF
                            UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                        Case Else
                        DrawTheme
                    End Select
                    If Not .UseTheme Then
                        'Draw Failed, use Classic Style
                        UserControl.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbHighlight, B
                    End If
                    'Add back image
                    Dim dx As Integer, dy As Integer
                    'On Error Resume Next
                    If Not m_SpecialGroupBackground Is Nothing Then
                        dx = m_SpecialGroupBackground.Width / Screen.TwipsPerPixelX
                        dy = m_SpecialGroupBackground.Height / Screen.TwipsPerPixelY
                        UserControl.ScaleMode = 3
                        UserControl.PaintPicture m_SpecialGroupBackground, .Left + 1, .Top + 1, .Width - 2, .Height - 2, , , , , vbSrcAnd
                    End If
                    'AlphaPaintPicture .Left + 1, .Top + 1, .Width - 2, .Height - 2, m_SpecialGroupBackground, 32
                    'Draw the items
                    For nj = 1 To m_SpecialGroup.iItemsCount
                        RedrawItem -1, nj, 0
                    Next nj
                    'group has been drawn
                    iTop = iTop + 6
                End If
                iTop = iTop + 6
            End If
            'for each group:
            For ni = 1 To iGroups
                'Draw Header
                RedrawGroupHeader ni
                If cGroups(ni).bExpanded Then
                    If Not cGroups(ni).pChild Is Nothing Then
                        'Child Picture Box Is Defined!
                        On Error Resume Next
                        cGroups(ni).pChild.Move cGroups(ni).mRect.Left * Screen.TwipsPerPixelX, (cGroups(ni).mRect.Bottom) * Screen.TwipsPerPixelY, (cGroups(ni).mRect.Right - cGroups(ni).mRect.Left) * Screen.TwipsPerPixelX
                        cGroups(ni).pChild.Visible = True
                        .hdc = cGroups(ni).pChild.hdc
                        .hwnd = cGroups(ni).pChild.hwnd
                        .Left = 0: .Top = 0: .Width = cGroups(ni).pChild.ScaleWidth: .Height = cGroups(ni).pChild.ScaleHeight
                        .Part = 5
                        .State = 1
                        cGroups(ni).pChild.AutoRedraw = True
                        'Draw the group Items frame
                        .Text = ""
                        .Part = 5
                        .State = 1
                        m_pChild_Paint (ni)
                        'DrawTheme
                        .hdc = UserControl.hdc
                        .hwnd = UserControl.hwnd
                    Else
                        'Draw the group Items frame
                        .Top = cGroups(ni).mRect.Bottom
                        .Left = cGroups(ni).mRect.Left
                        .Height = cGroups(ni).lItemsHeight
                        .Width = cGroups(ni).mRect.Right - cGroups(ni).mRect.Left
                        .Text = ""
                        .Part = 5
                        .State = 1
                        Select Case sColorName
                        'this styles now are EMULATED. (just like microsoft does)
                            Case "Metallic"
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                            Case "HomeStead"
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                            Case Else
                            DrawTheme
                        End Select
                        If Not .UseTheme Then
                            'Draw Failed, use Classic Style
                            UserControl.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbButtonFace, B
                        End If
                        'Draw the items
                        For nj = 1 To cGroups(ni).iItemsCount
                            RedrawItem ni, nj, 0
                        Next nj
                    'group has been drawn
                    End If
                Else
                    'hide everything!
                    If Not cGroups(ni).pChild Is Nothing Then
                        'Child Picture Box Is Defined!
                        cGroups(ni).pChild.Visible = False
                    End If
                End If
            Next ni
            'Details Group
            If m_bDetailsGroup Then
                ' Draw The Details Header
                RedrawDetailsHeader
                If m_DetailsGroup.bExpanded Then
                    'Draw the Tittle and text
                    .Part = 5
                    .State = 1
                    .Top = m_DetailsGroup.mRect.Bottom
                    .Left = m_DetailsGroup.mRect.Left
                    .Height = m_DetailsGroup.lItemsHeight
                    .Width = m_DetailsGroup.mRect.Right - m_DetailsGroup.mRect.Left
                    .Text = ""
                        Select Case sColorName
                        'this styles now are EMULATED. (just like microsoft does)
                            Case "Metallic"
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF0, &HF1, &HF5), BF
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                            Case "HomeStead"
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), RGB(&HF6, &HF6, &HEC), BF 'GetSysColor(COLOR_HIGHLIGHTTEXT), BF
                                UserControl.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height), vbWhite, B
                            Case Else
                            DrawTheme
                        End Select
                        If Not .UseTheme Then
                            'Draw Failed, use Classic Style
                            UserControl.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), vbButtonFace, B
                        End If
                    ''On Error Resume Next
                    'There Is a Image?
                    If m_DetailsPicture Is Nothing Then
                        'No Image
                        'Draw Tittle
                        UserControl.FontUnderline = False
                        UserControl.FontBold = True
                        SetRect tmpRect, m_DetailsRect.Left, m_DetailsGroup.mRect.Bottom + 11, UserControl.ScaleWidth - 32, m_DetailsGroup.mRect.Bottom + 68
                        DrawText UserControl.hdc, m_DetailsGroupTittle, -1, tmpRect, DT_LEFT Or DT_WORDBREAK
                        'DrawText
                        UserControl.FontBold = False
                        DrawText UserControl.hdc, m_DetailsGroupText, -1, m_DetailsRect, DT_LEFT Or DT_WORDBREAK 'Len(m_DetailsGroupText)
                        RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
                        'group has been drawn
                    Else
                        'We Have an Image move rects and go on
                        Dim lx As Integer, ly As Integer
                        If m_DetailsPicture.Width > m_DetailsPicture.Height Then
                            'Calculate size
                            'wip
                        Else
                            'Calculate size again o_0
                            'wip
                        End If
                        'Draw Tittle
                        UserControl.FontUnderline = False
                        UserControl.FontBold = True
                        SetRect tmpRect, m_DetailsRect.Left, m_DetailsGroup.mRect.Bottom + 11 + UserControl.ScaleWidth - 128, UserControl.ScaleWidth - 128, m_DetailsGroup.mRect.Bottom + 11 + UserControl.ScaleWidth
                        DrawText UserControl.hdc, m_DetailsGroupTittle, -1, tmpRect, DT_LEFT Or DT_WORDBREAK
                        'DrawText
                        UserControl.FontBold = False
                        DrawText UserControl.hdc, m_DetailsGroupText, -1, m_DetailsRect, DT_LEFT Or DT_WORDBREAK 'Len(m_DetailsGroupText)
                        RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
                        'Draw Image
                        UserControl.PaintPicture m_DetailsPicture, _
                         64, .Top + 8, UserControl.ScaleWidth - 128, UserControl.ScaleWidth - 128
                        'Draw Tittle
                    'group has been drawn
                    End If
                End If
                iTop = iTop + 20
            End If
            
        End With
    End If
    UserControl.Refresh
End Sub

'*************************************************************
'
'   Private Functions
'
'   Required Functions to make easier this ..thing
'
'**************************************

' Desc: On Version 0.9 and previous the rects of each item
'       where calulated on the Paint event of the usercontrol.
'       It Generated some problems, So I moved all that code
'       to a New Function. I Earned then almost 100 lines of
'       code!
Private Sub CalcRects()
    Dim ni As Integer, nj As Integer
    Dim iTop As Integer
    Dim bUseTheme As Boolean
    Dim itemRect As Rect
    Dim ItemWidth As Long
    'Start variables
    iTop = -m_iTopOffset
    UserControl.FontBold = False
    'iTop = -m_iTopOffset
    m_Width = IIf(m_ScrollBar.Visible, UserControl.ScaleWidth - m_ScrollBar.Width, UserControl.ScaleWidth)
    'Check for the Special Group
    If m_bSpecialGroup Then
        'Set properties for the Special group
        iTop = iTop + 16    'Top Offset
        m_SpecialGroup.mRect.Top = iTop
        m_SpecialGroup.mRect.Left = 8
        m_SpecialGroup.mRect.Bottom = iTop + 24
        m_SpecialGroup.mRect.Right = m_Width - 8
        iTop = m_SpecialGroup.mRect.Bottom
        If m_SpecialGroup.bExpanded Then
            'Calculate Item's Rects
            iTop = iTop + 10
            For nj = 1 To m_SpecialGroup.iItemsCount
                m_SpecialGroup.items(nj).mRect.Top = iTop
                m_SpecialGroup.items(nj).mRect.Left = 20
                m_SpecialGroup.items(nj).mRect.Right = 40 + IIf(TextWidth((m_SpecialGroup.items(nj).Caption)) + 1 < (m_Width - 56), TextWidth((m_SpecialGroup.items(nj).Caption)) + 1, m_Width - 56)
                m_SpecialGroup.items(nj).mRect.Bottom = iTop + CalcHeightRectText(40, m_Width - 16, m_SpecialGroup.items(nj).Caption)
                iTop = m_SpecialGroup.items(nj).mRect.Bottom + 8
            Next nj
            m_SpecialGroup.lItemsHeight = iTop - m_SpecialGroup.mRect.Bottom + 8
            'group has been calculated
            iTop = iTop + 6
        End If
        iTop = iTop + 6
    End If
    'for each group:
    For ni = 1 To iGroups
        'Calc Header Rect
        iTop = iTop + 10
        'Get Coordinates
        cGroups(ni).mRect.Top = iTop
        cGroups(ni).mRect.Left = 8
        cGroups(ni).mRect.Bottom = iTop + 24
        cGroups(ni).mRect.Right = m_Width - 8
        iTop = iTop + 24
        If cGroups(ni).bExpanded Then
            If Not cGroups(ni).pChild Is Nothing Then
                'Child Picture Box Is Defined!
                On Error Resume Next
                'Calculate the group Height
                iTop = iTop + cGroups(ni).pChild.ScaleHeight
                cGroups(ni).lItemsHeight = cGroups(ni).pChild.ScaleHeight
                'group has been Calculated
                iTop = iTop - 10
            Else
                iTop = iTop + 10
                'Calc the items
                For nj = 1 To cGroups(ni).iItemsCount
                    cGroups(ni).items(nj).mRect.Top = iTop
                    cGroups(ni).items(nj).mRect.Left = 20
                    cGroups(ni).items(nj).mRect.Right = 40 + IIf(TextWidth((cGroups(ni).items(nj).Caption)) + 1 < (m_Width - 56), TextWidth((cGroups(ni).items(nj).Caption)) + 1, m_Width - 56)
                    cGroups(ni).items(nj).mRect.Bottom = iTop + CalcHeightRectText(40, m_Width - 16, cGroups(ni).items(nj).Caption)
                
                    iTop = cGroups(ni).items(nj).mRect.Bottom + 8
                Next nj
                'Calculate the group Items frame
                cGroups(ni).lItemsHeight = iTop - cGroups(ni).mRect.Bottom + 12
                'group has been Calculated
                iTop = iTop + 6
            End If
        End If
        iTop = iTop + 12
    Next ni
    'Details Group
    If m_bDetailsGroup Then
        iTop = iTop + 8
        'Get Coordinates
        m_DetailsGroup.mRect.Top = iTop
        m_DetailsGroup.mRect.Left = 8
        m_DetailsGroup.mRect.Bottom = iTop + 24
        m_DetailsGroup.mRect.Right = m_Width - 8
        iTop = m_DetailsGroup.mRect.Bottom
        Dim iTittleHeight As Integer
        If m_DetailsGroup.bExpanded Then
            'If there Is a Details Image...
            On Error Resume Next
            If m_DetailsPicture Is Nothing Then
                'There Isn't a Image
                UserControl.FontBold = True
                iTittleHeight = CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupTittle)
                UserControl.FontBold = False
                m_DetailsGroup.lItemsHeight = iTittleHeight + CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupText) + 24
                'Set the Details Rect
                UserControl.FontBold = True
                SetRect m_DetailsRect, 20, iTop + CalcHeightRectText(20, m_Width - 24, m_DetailsGroupTittle) + 12, m_Width - 24, iTop + 20 + m_DetailsGroup.lItemsHeight
                UserControl.FontBold = False
                iTop = m_DetailsRect.Bottom '+ 4
            Else
                'We Have An Image make room for It.
                iTop = iTop + 12 + UserControl.ScaleWidth - 128
                'Calculate the pos of the text and the tittle
                'Get the Height of the text
                UserControl.FontBold = True
                iTittleHeight = CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupTittle)
                UserControl.FontBold = False
                m_DetailsGroup.lItemsHeight = iTittleHeight + CalcHeightRectText(20, UserControl.ScaleWidth - 32, m_DetailsGroupText) + 24
                'Set the Details Rect
                UserControl.FontBold = True
                SetRect m_DetailsRect, 20, iTop + CalcHeightRectText(20, m_Width - 24, m_DetailsGroupTittle) + 12, m_Width - 24, iTop + 20 + m_DetailsGroup.lItemsHeight
                UserControl.FontBold = False
                iTop = m_DetailsRect.Bottom '+ 4
                m_DetailsGroup.lItemsHeight = iTop - m_DetailsGroup.mRect.Bottom - 12
            End If
        'group has been drawn
        End If
    End If
    'I'm re-using this variable, sorry,  Idon't want more variables on this sub.
    'this var should be called something like ScrollAmount
    'anyway, I think nobody will read this stuff:P If you do, thanks for look
    'into this code. Check out the Rect's array for each item in each group, I liked It a Lot!
    ItemWidth = iTop - UserControl.ScaleHeight + m_iTopOffset
    If ItemWidth = 0 Then
        'Setup ScrollBar
        'Adjust ScrollBar Properties
        m_ScrollBar.SmallChange = 4
        m_ScrollBar.LargeChange = UserControl.ScaleHeight
        m_ScrollBar.Max = 1 '(-ItemWidth) - 40
        m_ScrollBar.Move UserControl.ScaleWidth - m_ScrollBar.Width, 0, m_ScrollBar.Width, UserControl.ScaleHeight
        If m_ScrollBar.Visible = True Then
            m_ScrollBar.Visible = False
            CalcRects
        End If
        SetRect m_RedrawRect, 1, 1, UserControl.ScaleWidth - m_ScrollBar.Width - 2, UserControl.ScaleHeight - 2
        m_iTopOffset = 0
    ElseIf ItemWidth < 0 Then
        'Hide ScrollBar
        If m_ScrollBar.Visible Then
            m_ScrollBar.Visible = False
            m_iTopOffset = 0
            CalcRects
            SetRect m_RedrawRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
            Exit Sub
        Else
            m_ScrollBar.Visible = False
        End If
    Else
        'show and update scrollbar
        On Error GoTo NoHeight
        m_ScrollBar.SmallChange = 4
        m_ScrollBar.LargeChange = UserControl.ScaleHeight
        m_ScrollBar.Max = ItemWidth
        m_ScrollBar.Move UserControl.ScaleWidth - m_ScrollBar.Width, 0, m_ScrollBar.Width, UserControl.ScaleHeight
        SetRect m_RedrawRect, 0, 0, UserControl.ScaleWidth - m_ScrollBar.Width - 1, UserControl.ScaleHeight
        If Not m_ScrollBar.Visible Then
            'Prevent Infinite loop
            If Not UserControl.Extender.Visible Then Exit Sub
            'Scrollbar was not visible, recalculate rects, but before set to visible.
            m_ScrollBar.Visible = True
            DoEvents
            If m_AllowRedraw Then
                ''Debug.Print "forced to calcrects!"
                CalcRects
            End If
        End If
    End If
Exit Sub
NoHeight:
    RaiseWarning "Couldn't Set ScrollBar Properties"
End Sub

' Desc: Calculate the height of a group box.
'       If there are multiline items, the
'       height won't be items*itemheight
Private Function CalcGroupHeight(iGroup As Integer) As Integer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nj As Integer, iTop As Integer
50020     Dim tmpHeight As Long           'USed to keep a copy oh m_LastTextHeight
50030     Dim textRect As Rect            'Copy used to calculate text height
50040     iTop = 24                       'Start up Offset
50050     tmpHeight = m_LastTextHeight    'Save var
50060     If iGroup = -1 Then
50070         With m_SpecialGroup
50080             For nj = 1 To .iItemsCount
50090                 '.items(nj).mRect.Top = iTop
50100                 SetRect textRect, _
                        .items(nj).mRect.Left + 20, _
                        .items(nj).mRect.Top, IIf((.items(nj).mRect.Right > m_Width - 12), _
                        m_Width - 12, .items(nj).mRect.Right), _
                        .items(nj).mRect.Bottom
50150                 m_LastTextHeight = CalcHeightRectText(textRect.Left, textRect.Right, .items(nj).Caption)
50160                 iTop = iTop + m_LastTextHeight + 8
50170             Next nj
50180         End With
50190         CalcGroupHeight = iTop
50200     Else    'Aplicar a grupo normal.
50210         With cGroups(iGroup)
50220             For nj = 1 To .iItemsCount
50230                 'textRect.Top = iTop    'I don't know why I wrote this :/ ( Now I know, It's for looping o_0
50240                 'Set the temp Rect
50250                 SetRect textRect, _
                        .items(nj).mRect.Left + 20, _
                        .items(nj).mRect.Top, IIf((.items(nj).mRect.Right > m_Width - 12), _
                        m_Width - 12, .items(nj).mRect.Right), _
                        .items(nj).mRect.Bottom
50300                 m_LastTextHeight = CalcHeightRectText(textRect.Left, textRect.Right, .items(nj).Caption)
50310                 iTop = iTop + m_LastTextHeight + 8
50320             Next nj
50330         End With
50340         CalcGroupHeight = iTop
50350         'CalcGroupHeight = cGroups(iGroup).iItemsCount * 24 + 10
50360     End If
50370     m_LastTextHeight = tmpHeight    'Restore Var
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "CalcGroupHeight")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Draw Multilined text.
' Returns: Height of drawed Text
Private Function DrawRectText(rtRect As Rect, sText As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'draw text in the selected position
50020     m_LastTextHeight = CalcHeightRectText(rtRect.Left, rtRect.Right, sText)
50030     rtRect.Bottom = rtRect.Top + m_LastTextHeight
'50040     DrawText UserControl.hdc, sText, Len(sText), rtRect, DT_LEFT Or DT_WORDBREAK
50040     DrawText UserControl.hdc, sText, -1, rtRect, DT_LEFT Or DT_WORDBREAK
50050     'Redraw Window
50060     RedrawWindow UserControl.hwnd, rtRect, ByVal 0&, RDW_INVALIDATE
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "DrawRectText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Draw Multilined text.
' Returns: Height of drawed Text
Private Function CalcHeightRectText(lLeft As Long, lright As Long, sText As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Calculate vertical height of text Tittle + Text(wrapped)
50020     Dim rectText As Rect
50030     SetRect rectText, lLeft, 0, lright, UserControl.ScaleHeight
'50040     CalcHeightRectText = DrawText(UserControl.hdc, sText, Len(sText), rectText, DT_CALCRECT Or DT_LEFT Or DT_WORDBREAK)
50040     CalcHeightRectText = DrawText(UserControl.hdc, sText, -1, rectText, DT_CALCRECT Or DT_LEFT Or DT_WORDBREAK)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "CalcHeightRectText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Determine If the mouse cursor is inside a Object
Private Function InBox(ObjectHWnd As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim mpos As POINT
50020     Dim oRect As Rect
50030     GetCursorPos mpos
50040     GetWindowRect ObjectHWnd, oRect
50050     If mpos.x >= oRect.Left And mpos.x <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
50070         InBox = True
50080     Else
50090         InBox = False
50100    End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "InBox")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub APILine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Use the API LineTo for Fast Drawing
50020     Dim pt As POINT
50030     UserControl.ForeColor = lColor
50040     MoveToEx UserControl.hdc, X1, Y1, pt
50050     LineTo UserControl.hdc, X2, Y2
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "APILine")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Make Soft a color
Function SoftColor(lColor As OLE_COLOR) As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim lRed As OLE_COLOR
50020     Dim lGreen As OLE_COLOR
50030     Dim lBlue As OLE_COLOR
50040     Dim lr As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
50050     lr = (lColor And &HFF)
50060     lg = ((lColor And 65280) \ 256)
50070     lb = ((lColor) And 16711680) \ 65536
50080     lRed = (76 - Int(((lColor And &HFF) + 32) \ 64) * 19)
50090     lGreen = (76 - Int((((lColor And 65280) \ 256) + 32) \ 64) * 19)
50100     lBlue = (76 - Int((((lColor And &HFF0000) \ &H10000) + 32) / 64) * 19)
50110     SoftColor = RGB(lr + lRed, lg + lGreen, lb + lBlue)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SoftColor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function TranslateColor(origincolor As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     TranslateColor = OleTranslateColor(origincolor, 0, 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "TranslateColor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Function DoGradient(FromColor As Long, ToColor As Long, Optional DrawHorVer As GRADIENT_FILL_RECT = FillHor, Optional Left As Long = 0, Optional Top As Long = 0, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim vert(1) As TRIVERTEX
50020     Dim gRect As GRADIENT_RECT
50030     Dim R As Byte, G As Byte, B As Byte
50040
50050     Long2RGB FromColor, R, G, B
50060     With vert(0)
50070         .x = Left
50080         .Y = Top
50090         .Red = Val("&h" & Hex(R) & "00")
50100         .Green = Val("&h" & Hex(G) & "00")
50110         .Blue = Val("&h" & Hex(B) & "00")
50120         .Alpha = 0&
50130     End With
50140
50150     Long2RGB ToColor, R, G, B
50160     With vert(1)
50170         .x = Left + Width
50180         .Y = Top + Height
50190         .Red = Val("&h" & Hex(R) & "00")
50200         .Green = Val("&h" & Hex(G) & "00")
50210         .Blue = Val("&h" & Hex(B) & "00")
50220         .Alpha = 0&
50230     End With
50240
50250     gRect.UpperLeft = 0
50260     gRect.LowerRight = 1
50270
50280     DoGradient = GradientFillRect(UserControl.hdc, vert(0), 2, gRect, 1, DrawHorVer)
50290
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "DoGradient")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Function Long2RGB(nColor As Long, Red As Byte, Green As Byte, Blue As Byte)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Red = (nColor And &HFF&)
50020     Green = (nColor And &HFF00&) / &H100
50030     Blue = (nColor And &HFF0000) / &H10000
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Long2RGB")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


'' Desc: a Alpha version of the paintpicture function
''       Still don't used
'Private Sub AlphaPaintPicture(ByVal x As Long, ByVal y As Long, ByVal lwidth As Long, ByVal lheight As Long, lPicture As Picture, Optional ByVal lConstantAlpha As Byte = 255)
''Heavily based on this post:
''http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43879&lngWId=1
''with the suggestion on Dana Seaman and edited by me to make It a useful function
'    Dim lr As Long
'    Dim BF As BLENDFUNCTION
'    Dim lBF As Long
'    'The Structure will be replaced.
'    With BF
'        .BlendOp = AC_SRC_OVER
'        .BlendFlags = 0
'        .SourceConstantAlpha = lConstantAlpha
'        .AlphaFormat = 0
'    End With
'    'copy the BLENDFUNCTION-structure to a Long
'    RtlMoveMemory lBF, BF, 4
'
'    lBF = &H10000 * lConstantAlpha
'    m_tempImg.ScaleMode = 3
'    m_tempImg.Width = lPicture.Width / Screen.TwipsPerPixelX
'    m_tempImg.Height = lPicture.Height / Screen.TwipsPerPixelY
'    Set m_tempImg.Picture = lPicture
'    Set frmTest.Picture5.Picture = m_tempImg.Picture
'    'AlphaBlend
'    lr = AlphaBlend(UserControl.hdc, x, y, lwidth, lheight, m_tempImg.hdc, 0, 0, m_tempImg.ScaleWidth, m_tempImg.ScaleHeight, lBF)
'    If (lr = 0) Then
'       RaiseWarning Err.LastDllError
'    End If
'
'End Sub

' Desc: Convert a RGB color to long
Private Function RGBToLong(rgbColor As RGB) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     RGBToLong = rgbColor.Blue + rgbColor.Green * 265 + rgbColor.Red * 65536
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "RGBToLong")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc Convert a long into a RGB structure
Private Function LongToRGB(lColor As Long) As RGB
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     LongToRGB.Red = lColor And &HFF
50020     LongToRGB.Green = (lColor \ &H100) And &HFF
50030     LongToRGB.Blue = (lColor \ &H10000) And &HFF
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "LongToRGB")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


' Desc: This function will return  whether you are running
'       your program or DLL from within the IDE, or compiled.
Private Function InVBDesignEnvironment() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 'Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11615&lngWId=1
50020     Dim strFileName As String
50030     Dim lngCount As Long
50040
50050     strFileName = String(255, 0)
50060     lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
50070     strFileName = Left(strFileName, lngCount)
50080
50090     InVBDesignEnvironment = False
50100
50110     If UCase(Right(strFileName, 7)) = "VB5.EXE" Then
50120         InVBDesignEnvironment = True
50130     ElseIf UCase(Right(strFileName, 7)) = "VB6.EXE" Then
50140         InVBDesignEnvironment = True
50150     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "InVBDesignEnvironment")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Get the Hand Cursor
Public Sub SetHandCur(Hand As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     If Hand = True Then
50020         UserControl.MousePointer = 99
50030     Else
50040         UserControl.MousePointer = 0
50050     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetHandCur")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Get a Group Index By Key
Private Function GetGroupsByKeyN(ByVal sGroupKey As Variant) As Integer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     Dim ni As Integer
50030
50040     If (VarType(sGroupKey) <> vbInteger) And (VarType(sGroupKey) <> vbString) Then
50050         RaiseError "GetGroupsByKeyN: sGroupKey not of required Type (String or Integer)!"
50060         GetGroupsByKeyN = -3
50070         Exit Function
50080     End If
50090     'KEY was passed?
50100     If VarType(sGroupKey) = vbString Then
50110         'Check Normal Groups
50120         For ni = 1 To iGroups
50130             If sGroupKey = cGroups(ni).key Then
50140                 'this is the index
50150                 GetGroupsByKeyN = ni
50160                 Exit Function
50170             End If
50180         Next ni
50190         'Check Special Group
50200         If sGroupKey = "Special Group" Then
50210             GetGroupsByKeyN = -1
50220             Exit Function
50230         'Check Details Group
50240         ElseIf sGroupKey = "Details" Then
50250             GetGroupsByKeyN = -2
50260             Exit Function
        'Finally: String didn't match
50280         Else
50290             GetGroupsByKeyN = -3
50300             Exit Function
50310         End If
50320     'INDEX was passed
50330     Else
50340         GetGroupsByKeyN = sGroupKey
50350     End If
50360
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "GetGroupsByKeyN")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Redraw a Single Item
Private Function RedrawItem(iCurrentGroup As Integer, iItemNum As Integer, iState As Integer)
    'Dim lTextColor As Long
    Dim textRect As Rect
    'Set the text color
    Select Case iState
        Case 1  'Normal
            UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
        Case 2  'Hover
            UserControl.ForeColor = GetSysColor(COLOR_HIGHLIGHT)
        Case 3  'hot
            UserControl.ForeColor = GetSysColor(COLOR_GRADIENTACTIVECAPTION)
        Case 4  'Disabled
            UserControl.ForeColor = GetSysColor(COLOR_GRAYTEXT)
    End Select
    'Use underline style
    UserControl.FontUnderline = True
    UserControl.FontBold = False
    If iCurrentGroup = -1 Then
        With m_SpecialGroup.items(iItemNum)
            'Check for multiline text
            'if multiline text, adjust right
            'and adjust left to make room for the image
            SetRect textRect, _
                    .mRect.Left + 20, _
                    .mRect.Top, _
                    m_Width - 12, _
                    .mRect.Bottom
            DrawRectText textRect, .Caption
            On Error GoTo NoImage
            If iImgLType = 1 Then
                UserControl.PaintPicture m_objImageList.ListImages(.Icon).ExtractIcon, .mRect.Left, .mRect.Top, 16, 16
            ElseIf iImgLType = 2 Then
                m_objImageList.DrawImage .Icon, UserControl.hdc, .mRect.Left, .mRect.Top
            End If
        End With
    Else
        With cGroups(iCurrentGroup).items(iItemNum)
            'Set the rect where the text will be drawn
            SetRect textRect, _
                    .mRect.Left + 20, _
                    .mRect.Top, _
                    m_Width - 12, _
                    .mRect.Bottom
            'Draw the text
            DrawRectText textRect, .Caption
            On Error GoTo NoImage
            'Try to Draw the item image
            If iImgLType = 1 Then
                UserControl.PaintPicture m_objImageList.ListImages(.Icon).ExtractIcon, .mRect.Left, .mRect.Top, 16, 16
            ElseIf iImgLType = 2 Then
                m_objImageList.DrawImage .Icon, UserControl.hdc, .mRect.Left, .mRect.Top
            End If
        End With
    End If
    UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
Exit Function
NoImage:
    'No image or not imagelist was selected
    RaiseWarning "No Defined Imagelist or invalid Image Index"
End Function

' Desc: Redraw a Group Header:
Private Function RedrawGroupHeader(iCurrentGroup As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim textRect As Rect
50020     Dim lcolor1 As Long, lcolor2 As Long
50030     'Setup Variables
50040     UserControl.FontUnderline = False
50050     UserControl.FontBold = True
50060     With cGroups(iCurrentGroup)
50070         m_cUxTheme.Part = 8
50080         m_cUxTheme.Left = .mRect.Left
50090         m_cUxTheme.Top = .mRect.Top
50100         m_cUxTheme.Width = .mRect.Right - .mRect.Left
50110         m_cUxTheme.Height = .mRect.Bottom - .mRect.Top
50120         m_cUxTheme.State = .iState 'Now Support More States
50130         m_cUxTheme.Text = cGroups(iCurrentGroup).Caption
50140         m_cUxTheme.TextOffset = 0
50150         'Search for current theme and color scheme
50160         'Microsoft created the ExplorerBar with custom code and Images.
50170         'So We Need do somethig Similar. we will search for the theme file
50180         'and color Scheme
50191         Select Case sColorName
                  Case "HomeStead"
50210                 'this styles now are EMULATED. (just like microsoft does)
50220                 DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
50230                 DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
50240                 DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
50250                 SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
50260                 UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_HIGHLIGHT), GetSysColor(COLOR_3DDKSHADOW))
50270                 UserControl.FontUnderline = False
50280                 UserControl.FontBold = True
50290                 DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
50300                 RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
50310                 UserControl.ForeColor = vbButtonText
50320                 UserControl.FontUnderline = True
50330                 UserControl.FontBold = False
50340             Case "Metallic"
50350                 'this styles now are EMULATED. (just like microsoft does)
50360                 DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
50370                 DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
50380                 DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
50390                 SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
50400                 UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_3DDKSHADOW), GetSysColor(COLOR_BTNTEXT))
50410                 UserControl.FontUnderline = False
50420                 UserControl.FontBold = True
50430                 DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
50440                 RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
50450                 UserControl.ForeColor = vbButtonText
50460                 UserControl.FontUnderline = True
50470                 UserControl.FontBold = False
50480             Case Else '"blue" and other themes
50490                 DrawTheme
50500         End Select
50510         If Not m_cUxTheme.UseTheme Then
50520             'no theme aviable, use classic style
50530             SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
50540             UserControl.Line (.mRect.Left, .mRect.Top)-(.mRect.Right, .mRect.Bottom), vbButtonFace, BF
50550             UserControl.ForeColor = vbButtonText
50560             UserControl.FontUnderline = False
50570             UserControl.FontBold = True
50580             DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
50590             RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
50600             UserControl.ForeColor = vbButtonText
50610             UserControl.FontUnderline = True
50620             UserControl.FontBold = False
50630         End If
50640
50650         'Draw Expand Button
50660         m_cUxTheme.Part = 7 + .bExpanded
50670         m_cUxTheme.Text = ""
50680         m_cUxTheme.Top = .mRect.Top
50690         m_cUxTheme.Left = m_Width - 32
50700         m_cUxTheme.Width = 24
50710         m_cUxTheme.Height = 24
50720         m_cUxTheme.State = .iState
50731         Select Case sColorName
                  'this styles now are EMULATED. (just like microsoft does)
                  Case "Metallic"
50760                 UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 0, 17, 17, vbSrcCopy
50770             Case "HomeStead"
50780                 UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 18, 17, 17, vbSrcCopy
50790             Case Else
50800             DrawTheme
50810         End Select
50820         If Not m_cUxTheme.UseTheme Then
50830             'no theme aviable, use classic style
50840             If .iState = 3 Then  'pressed
50850                 lcolor2 = vb3DHighlight: lcolor1 = vb3DShadow
50860             ElseIf .iState = 2 Then 'Hover
50870                 lcolor1 = vb3DHighlight: lcolor2 = vb3DShadow
50880             Else    'Normal
50890                 lcolor1 = vbButtonFace: lcolor2 = vbButtonFace
50900             End If
50910             'Draw Dutton
50920             UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
50930             UserControl.Line (m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
50940             UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor1
50950             UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4), lcolor1
50960             'Draw arrow
50970             DrawArrow m_cUxTheme.Left, m_cUxTheme.Top, .bExpanded, vbButtonText
50980
50990         End If
51000     RedrawWindow UserControl.hwnd, .mRect, ByVal 0&, RDW_INVALIDATE
51010     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "RedrawGroupHeader")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Redraw a special Group Header
Private Function RedrawSpecialHeader()
    Dim lcolor1 As Long, lcolor2 As Long
    Dim textRect As Rect
    On Error Resume Next
    UserControl.FontUnderline = False
    UserControl.FontBold = True
    With m_SpecialGroup
        m_cUxTheme.Part = 12
        m_cUxTheme.Left = .mRect.Left
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Width = .mRect.Right - .mRect.Left
        m_cUxTheme.Height = .mRect.Bottom - .mRect.Top
        m_cUxTheme.State = .iState '(Doesn't support other states )
        m_cUxTheme.Text = .Caption
        m_cUxTheme.TextOffset = 36
        Select Case sColorName
            Case "Metallic"
                'this styles now are EMULATED. (just like microsoft does)
                DoGradient RGB(&H77, &H77, &H92), RGB(&HB4, &HB6, &HC7), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
                DoGradient RGB(&H77, &H77, &H92), RGB(&HB4, &HB6, &HC7), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
                DoGradient RGB(&H77, &H77, &H92), RGB(&HB4, &HB6, &HC7), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
                SetRect textRect, .mRect.Left + m_cUxTheme.TextOffset, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
                UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_BTNFACE), GetSysColor(COLOR_BTNHIGHLIGHT))
                UserControl.FontUnderline = False
                UserControl.FontBold = True
                DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
                RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
                UserControl.ForeColor = vbButtonText
                UserControl.FontUnderline = True
                UserControl.FontBold = False
            Case Else
                DrawTheme
        End Select
        If Not m_cUxTheme.UseTheme Then
            'no theme aviable, use classic style
            SetRect textRect, .mRect.Left + m_cUxTheme.TextOffset, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
            UserControl.Line (.mRect.Left, .mRect.Top)-(.mRect.Right, .mRect.Bottom), vbHighlight, BF
            UserControl.ForeColor = vbHighlightText
            UserControl.FontUnderline = False
            UserControl.FontBold = True
            DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
            RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
            UserControl.ForeColor = vbBlack
            UserControl.FontUnderline = True
            UserControl.FontBold = False
        End If
        'm_cUxTheme.DrawThemeTextEx 1, iState
        'Draw Expand Button
        m_cUxTheme.TextOffset = 0
        m_cUxTheme.Part = 11 + .bExpanded
        m_cUxTheme.Text = ""
        m_cUxTheme.Top = .mRect.Top
        m_cUxTheme.Left = m_Width - 32
        m_cUxTheme.Width = 24
        m_cUxTheme.Height = 24
        m_cUxTheme.State = .iState
        Select Case sColorName
            Case "Metallic"
                UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 17 * -.bExpanded, 0, 17, 17, vbSrcCopy
            Case Else
                DrawTheme
        End Select
        If Not m_cUxTheme.UseTheme Then
            'no theme aviable, use classic style
            If .iState = 3 Then  'Pressed
                lcolor2 = vb3DHighlight: lcolor1 = vb3DShadow
            ElseIf .iState = 2 Then 'Hover
                lcolor1 = vb3DHighlight: lcolor2 = vb3DShadow
            Else    'normal
                lcolor1 = vbHighlight: lcolor2 = vbHighlight
            End If
            'Draw Dutton
            UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
            UserControl.Line (m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
            UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor1
            UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4), lcolor1
            'Draw arrow
            DrawArrow m_cUxTheme.Left, m_cUxTheme.Top, .bExpanded, vbWindowBackground
        End If
    UserControl.PaintPicture m_SpecialGroupIcon, 12, .mRect.Top - 8, 32, 32 ', 0, 0, 32, 32
    RedrawWindow UserControl.hwnd, .mRect, ByVal 0&, RDW_INVALIDATE
    'm_LastTextHeight = .mRect.Bottom - .mRect.Top
    End With
End Function

' Desc: Redraw the Details Group Header
Private Function RedrawDetailsHeader()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim textRect As Rect
50020     Dim lcolor1 As Long, lcolor2 As Long
50030     UserControl.FontUnderline = False
50040     UserControl.FontBold = True
50050     With m_DetailsGroup
50060         m_cUxTheme.Part = 8
50070         m_cUxTheme.Left = .mRect.Left
50080         m_cUxTheme.Top = .mRect.Top
50090         m_cUxTheme.Width = .mRect.Right - .mRect.Left
50100         m_cUxTheme.Height = .mRect.Bottom - .mRect.Top
50110         m_cUxTheme.State = .iState '(Doesn't support other states )
50120         m_cUxTheme.Text = m_DetailsGroup.Caption
50130         m_cUxTheme.TextOffset = 0
50140         'Search for current theme and color scheme
50150         'Microsoft created the ExplorerBar with custom code and Images.
50160         'So We Need do somethig Similar. we will search for the theme file
50170         'and color Scheme
50181         Select Case sColorName
                  Case "HomeStead"
50200                 'this styles now are EMULATED. (just like microsoft does)
50210                 DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
50220                 DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
50230                 DoGradient RGB(&HFF, &HFC, &HEC), RGB(&HE0, &HE7, &HB8), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
50240                 SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
50250                 UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_HIGHLIGHT), GetSysColor(COLOR_3DDKSHADOW))
50260                 UserControl.FontUnderline = False
50270                 UserControl.FontBold = True
50280                 DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
50290                 RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
50300                 UserControl.ForeColor = vbButtonText
50310                 UserControl.FontUnderline = True
50320                 UserControl.FontBold = False
50330             Case "Metallic"
50340                 'this styles now are EMULATED. (just like microsoft does)
50350                 DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 2, .mRect.Top, .mRect.Right - .mRect.Left - 4, .mRect.Bottom - .mRect.Top
50360                 DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left + 1, .mRect.Top + 1, .mRect.Right - .mRect.Left - 2, .mRect.Bottom - .mRect.Top - 1
50370                 DoGradient vbWhite, RGB(&HD6, &HD7, &HE0), FillHor, .mRect.Left, .mRect.Top + 2, .mRect.Right - .mRect.Left, .mRect.Bottom - .mRect.Top - 2
50380                 SetRect textRect, .mRect.Left + 12, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
50390                 UserControl.ForeColor = IIf(.bOver, GetSysColor(COLOR_3DDKSHADOW), GetSysColor(COLOR_BTNTEXT))
50400                 UserControl.FontUnderline = False
50410                 UserControl.FontBold = True
50420                 DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
50430                 RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
50440                 UserControl.ForeColor = vbButtonText
50450                 UserControl.FontUnderline = True
50460                 UserControl.FontBold = False
50470             Case Else '"blue" and other themes
50480                 DrawTheme
50490         End Select
50500         If Not m_cUxTheme.UseTheme Then
50510             'no theme aviable, use classic style
50520             SetRect textRect, .mRect.Left + 4, .mRect.Top, .mRect.Right - 25, .mRect.Bottom
50530             UserControl.Line (.mRect.Left, .mRect.Top)-(.mRect.Right, .mRect.Bottom), vbButtonFace, BF
50540             UserControl.ForeColor = vbButtonText
50550             UserControl.FontUnderline = False
50560             UserControl.FontBold = True
50570             DrawText UserControl.hdc, .Caption, -1, textRect, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_MODIFYSTRING Or DT_WORD_ELLIPSIS
50580             RedrawWindow UserControl.hwnd, m_DetailsRect, ByVal 0&, RDW_INVALIDATE
50590             UserControl.ForeColor = vbButtonText
50600             UserControl.FontUnderline = True
50610             UserControl.FontBold = False
50620         End If
50630         'Draw Expand Button
50640         m_cUxTheme.Part = 7 + .bExpanded
50650         m_cUxTheme.State = .iState
50660         m_cUxTheme.Text = ""
50670         m_cUxTheme.Top = .mRect.Top
50680         m_cUxTheme.Left = m_Width - 32
50690         m_cUxTheme.Width = 24
50700         m_cUxTheme.Height = 24
50711         Select Case sColorName
                  'this styles now are EMULATED. (just like microsoft does)
                  Case "Metallic"
50740                 UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 0, 17, 17, vbSrcCopy
50750             Case "HomeStead"
50760                 UserControl.PaintPicture UserControl.imgbuttons.Picture, m_cUxTheme.Left + 4, m_cUxTheme.Top + 4, 17, 17, 34 + 17 * -.bExpanded, 18, 17, 17, vbSrcCopy
50770             Case Else
50780             DrawTheme
50790         End Select
50800         If Not m_cUxTheme.UseTheme Then
50810             'no theme aviable, use classic style
50820             If .iState = 3 Then  'Pressed
50830                 lcolor2 = vb3DHighlight: lcolor1 = vb3DShadow
50840             ElseIf .iState = 2 Then 'Hover
50850                 lcolor1 = vb3DHighlight: lcolor2 = vb3DShadow
50860             Else    'Normal
50870                 lcolor1 = vbButtonFace: lcolor2 = vbButtonFace
50880             End If
50890             'Draw Dutton
50900             UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
50910             UserControl.Line (m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor2
50920             UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + 4, m_cUxTheme.Top + m_cUxTheme.Height - 4), lcolor1
50930             UserControl.Line (m_cUxTheme.Left + 4, m_cUxTheme.Top + 4)-(m_cUxTheme.Left + m_cUxTheme.Width - 4, m_cUxTheme.Top + 4), lcolor1
50940             'Draw arrow
50950             DrawArrow m_cUxTheme.Left, m_cUxTheme.Top, .bExpanded, vbButtonText
50960         End If
50970         RedrawWindow UserControl.hwnd, .mRect, ByVal 0&, RDW_INVALIDATE
50980     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "RedrawDetailsHeader")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: Draw the selected theme class, part, state on the especified rect
Private Function DrawTheme() As Boolean
Dim hTheme As Long
Dim bSuccess As Boolean
Dim lr As Long
Dim tTextR As Rect
Dim tContentR As Rect
Dim tImlR As Rect
On Error Resume Next
With m_cUxTheme
    If sColorName = "Classic" Then
        .UseTheme = False
        DrawTheme = False
        Exit Function
    End If
   bSuccess = True
   hTheme = OpenThemeData(.hwnd, StrPtr(.sClass))
   If (hTheme) Then
      'We Got an htheme
      .UseTheme = True
      Dim tR As Rect
      Dim lWidthTaken As Long
      tR.Left = .Left
      tR.Top = .Top
      If (.IconIndex > -1) And (.hIml) Then
         ImageList_GetImageRect .hIml, .IconIndex, tImlR
         lWidthTaken = tImlR.Right - tImlR.Left + 4 + .TextOffset
      End If
      lWidthTaken = lWidthTaken + .TextOffset
      If (.UseThemeSize) Then
         Dim tSize As Size
         lr = GetThemePartSize(hTheme, .hdc, .Part, .State, tR, TS_TRUE, tSize)
         tR.Right = tR.Left + tSize.cx
         tR.Bottom = tR.Top + tSize.cy
         lr = GetThemeBackgroundContentRect(hTheme, .hdc, .Part, .State, tR, tContentR)
         If (.IconIndex > -1) And (.hIml) Then
            If ((tContentR.Bottom - tContentR.Top) < (tImlR.Bottom - tImlR.Top + 4)) Then
               tR.Bottom = tR.Bottom + ((tImlR.Bottom - tImlR.Top + 4) - (tContentR.Bottom - tContentR.Top))
            End If
            If ((tContentR.Right - tContentR.Left) < (tImlR.Right - tImlR.Left + 4)) Then
               tR.Right = tR.Right + ((tImlR.Right - tImlR.Left + 4) - (tContentR.Right - tContentR.Left))
            End If
         End If
         If Len(.Text) > 0 Then
            lr = GetThemeBackgroundContentRect(hTheme, .hdc, .Part, .State, tR, tContentR)
            lr = GetThemeTextExtent(hTheme, .hdc, .Part, .State, StrPtr(.Text), -1, .TextAlign, tR, tTextR)
            If ((tContentR.Bottom - tContentR.Top) < (tTextR.Bottom - tTextR.Top)) Then
               tR.Bottom = tR.Bottom + ((tTextR.Bottom - tTextR.Top) - (tContentR.Bottom - tContentR.Top))
            End If
            If ((tContentR.Right - tContentR.Left - lWidthTaken) < (tTextR.Right - tTextR.Left + 8)) Then
               tR.Right = tR.Right + ((tTextR.Right - tTextR.Left + 8) - (tContentR.Right - tContentR.Left - lWidthTaken))
            End If
         End If
      Else
         tR.Right = .Left + .Width
         tR.Bottom = .Top + .Height
      End If
      
      lr = DrawThemeParentBackground( _
         .hwnd, _
         .hdc, _
         tR)
      If (lr <> S_OK) Then
         bSuccess = False
         RaiseWarning "Failed to parent draw background for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
      End If
      lr = DrawThemeBackground( _
         hTheme, _
         .hdc, _
         .Part, _
         .State, _
         tR, tR)
      If (lr <> S_OK) Then
         bSuccess = False
         'Important this is the main theme drawing procedure,
         'If this fail, then we can say the entire sub has
         'failed.
        .UseTheme = False
         RaiseWarning "Failed to draw background for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
      End If
      If Len(.Text) > 0 Then
         lr = GetThemeBackgroundContentRect( _
            hTheme, _
            .hdc, _
            .Part, _
            .State, _
            tR, _
            tTextR)
         If (lr <> S_OK) Then
            bSuccess = False
            'RaiseWarning "Failed to retrieve background content rectangle for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
         End If
         tTextR.Left = tTextR.Left + lWidthTaken
         tTextR.Right = tR.Right - .RightTextOffset
         tTextR.Top = tR.Top
         tTextR.Bottom = tR.Bottom
         If UxThemeText Then
            'This will fail with far asian languages, replaced With custom DrawText
            lr = DrawThemeText( _
               hTheme, _
               .hdc, _
               .Part, _
               .State, _
                StrPtr(.Text), _
               -1, _
               .TextAlign, _
               0, _
               tTextR)
            Else
                Dim ltmpColor As Long
                ltmpColor = UserControl.ForeColor
                If .Part = 12 Then
                    UserControl.ForeColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
                Else
                    UserControl.ForeColor = IIf(.State = 1, GetSysColor(COLOR_HIGHLIGHT), SoftColor(GetSysColor(COLOR_HIGHLIGHT)))
                End If
                DrawText .hdc, .Text, -1, tTextR, .TextAlign
                UserControl.ForeColor = GetSysColor(COLOR_BTNTEXT)
            End If
         If (lr <> S_OK) Then
            bSuccess = False
            'RaiseWarning "Failed to draw theme text for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
         End If
      End If
      If (.IconIndex > -1) Then
         Dim tIconR As Rect
         lr = GetThemeBackgroundContentRect( _
            hTheme, _
            .hdc, _
            .Part, _
            .State, _
            tR, _
            tIconR)
         ImageList_GetImageRect .hIml, .IconIndex, tImlR
         tIconR.Left = tIconR.Left + 2
         tIconR.Top = tIconR.Top + 2
         tIconR.Right = tIconR.Left + tImlR.Right - tImlR.Left
         tIconR.Bottom = tIconR.Top + tImlR.Bottom - tImlR.Top
         lr = DrawThemeIcon( _
            hTheme, _
            .hdc, _
            .Part, _
            .State, _
            tIconR, _
            .hIml, _
            .IconIndex)
         If (lr <> S_OK) Then
            bSuccess = False
            'RaiseWarning "Failed to draw theme icon for class '" & .sClass & "', partId=" & .Part & ", stateId=" & .State
         End If
      End If
      CloseThemeData hTheme
      Dim tmpRect As Rect
      SetRect tmpRect, .Left, .Top, .Left + .Width, .Top + .Height
      RedrawWindow .hwnd, tmpRect, ByVal 0&, RDW_INVALIDATE
   Else
      RaiseWarning "No theme data for class '" & .sClass & "'.  - " & Err.LastDllError
      bSuccess = False
      .UseTheme = False
   End If
End With
   DrawTheme = bSuccess
End Function

Private Sub GetThemeName()
    'Gett the current Theme name, ans Scheme Color
    Dim hTheme As Long
    Dim sShellStyle As String
    Dim lPtrThemeFile As Long, lPtrColorName As Long, hres As Long
    Dim iPos As Long
    On Error Resume Next
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr("ExplorerBar"))
   
   If Not hTheme = 0 Then
      ReDim bThemeFile(0 To 260 * 2) As Byte
      lPtrThemeFile = VarPtr(bThemeFile(0))
      ReDim bColorName(0 To 260 * 2) As Byte
      lPtrColorName = VarPtr(bColorName(0))
      hres = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
      
      sThemeFile = bThemeFile
      iPos = InStr(sThemeFile, vbNullChar)
      If (iPos > 1) Then sThemeFile = Left(sThemeFile, iPos - 1)
      sColorName = bColorName
      iPos = InStr(sColorName, vbNullChar)
      If (iPos > 1) Then sColorName = Left(sColorName, iPos - 1)
      
      sShellStyle = sThemeFile
      For iPos = Len(sThemeFile) To 1 Step -1
         If (Mid(sThemeFile, iPos, 1) = "\") Then
            sShellStyle = Left(sThemeFile, iPos)
            Exit For
         End If
      Next iPos
      sShellStyle = sShellStyle & "Shell\" & sColorName & "\ShellStyle.dll"
      CloseThemeData hTheme
    Else
        sColorName = "Classic"
    End If

End Sub

' Desc: This small sub draws the arrow in the selected position
Private Sub DrawArrow(ByVal x As Integer, ByVal Y As Integer, ByVal bUp As Boolean, ByVal lColor As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     If bUp Then
50020         UserControl.Line (x + 9, Y + 11)-(x + 13, Y + 7), lColor
50030         UserControl.Line (x + 10, Y + 11)-(x + 13, Y + 8), lColor
50040         UserControl.Line (x + 15, Y + 11)-(x + 11, Y + 7), lColor
50050         UserControl.Line (x + 14, Y + 11)-(x + 11, Y + 8), lColor
50060         UserControl.Line (x + 9, Y + 15)-(x + 13, Y + 11), lColor
50070         UserControl.Line (x + 10, Y + 15)-(x + 13, Y + 12), lColor
50080         UserControl.Line (x + 15, Y + 15)-(x + 11, Y + 11), lColor
50090         UserControl.Line (x + 14, Y + 15)-(x + 11, Y + 12), lColor
50100     Else
50110         UserControl.Line (x + 9, Y + 8)-(x + 13, Y + 12), lColor
50120         UserControl.Line (x + 10, Y + 8)-(x + 13, Y + 11), lColor
50130         UserControl.Line (x + 15, Y + 8)-(x + 11, Y + 12), lColor
50140         UserControl.Line (x + 14, Y + 8)-(x + 11, Y + 11), lColor
50150         UserControl.Line (x + 9, Y + 12)-(x + 13, Y + 16), lColor
50160         UserControl.Line (x + 10, Y + 12)-(x + 13, Y + 15), lColor
50170         UserControl.Line (x + 15, Y + 12)-(x + 11, Y + 16), lColor
50180         UserControl.Line (x + 14, Y + 12)-(x + 11, Y + 15), lColor
50190     End If
50200
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "DrawArrow")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Show an Error to the programmer Is using the control
Private Sub RaiseError(sErrorDescription As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     MsgBox "An Error has ocurred!" & vbCrLf & _
            sErrorDescription, vbCritical, "isExplorerBar"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "RaiseError")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Show Warning in the Debug Window
Private Sub RaiseWarning(sWarning As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 '    Debug.Print "'*************************"
50020 '    Debug.Print "'*     isExplorer Warning."
50030 '    Debug.Print "'*     " & sWarning
50040 '    Debug.Print "'*************************"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "RaiseWarning")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Create tooltip!
' This is still a work in progress function
Private Function CreateTooltip(sTittle, sCaption) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim lpRect As Rect
50020     Dim lWinStyle As Long
50030
50040 '    If lHwnd <> 0 Then
50050 '        DestroyWindow lHwnd
50060 '    End If
50070
50080     lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
50090
50100     ''create baloon style if desired
50110     'If mvarStyle = TTBalloon Then
50120     lWinStyle = lWinStyle Or TTS_BALLOON
50130
50140     ''the parent control has to have been set first
50150     'If Not mvarParentControl Is Nothing Then
50160         m_ttlHwnd = CreateWindowEx(0&, _
                    TOOLTIPS_CLASSA, _
                    vbNullString, _
                    lWinStyle, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    UserControl.hwnd, _
                    0&, _
                    App.hInstance, _
                    0&)
50280
50290         ''make our tooltip window a topmost window
50300         SetWindowPos m_ttlHwnd, _
            HWND_TOPMOST, _
            0&, _
            0&, _
            0&, _
            0&, _
            SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
50370
50380         ''get the rect of the parent control
50390         GetClientRect UserControl.hwnd, lpRect
50400
50410         ''now set our tooltip info structure
50420         With m_tti
50430             ''if we want it centered, then set that flag
50440             'If mvarCentered Then
50450                 .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
50460             'Else
50470                 .lFlags = TTF_SUBCLASS
50480             'End If
50490
50500             ''set the hwnd prop to our parent control's hwnd
50510             .lHwnd = UserControl.hwnd
50520             .lId = 0
50530             .hInstance = App.hInstance
50540             '.lpstr = ALREADY SET
50550             .lpRect = lpRect
50560         End With
50570
50580         ''add the tooltip structure
50590         SendMessage m_ttlHwnd, TTM_ADDTOOLA, 0&, m_tti
50600
50610         ''if we want a title or we want an icon
50620         If m_ttTitle <> vbNullString Or m_ttIcon <> TTNoIcon Then
50630             SendMessage m_ttlHwnd, TTM_SETTITLE, CLng(m_ttIcon), ByVal m_ttTitle
50640         End If
50650
50660         If m_ttForeColor <> Empty Then
50670             SendMessage m_ttlHwnd, TTM_SETTIPTEXTCOLOR, m_ttForeColor, 0&
50680         End If
50690
50700         If m_ttBackColor <> Empty Then
50710             SendMessage m_ttlHwnd, TTM_SETTIPBKCOLOR, m_ttBackColor, 0&
50720         End If
50730
50740     'End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "CreateTooltip")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'**************************************************************
'Function:      GetItemIndex / Private
'Description:   Returns the index of an BarItem, passed as variant
'               Item can only be String or Integer here
'Parameters:    selGroup:   BarGroup, the group we search in
'               Item:       Variant, containing Items parameter
'Result:        0 for no item found
'               Items index for success
'**************************************************************

Private Function GetItemIndex(selgroup As BarGroup, Item As Variant) As Integer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     Dim nj As Integer
50030
50040     'Check if there are Items in the group
50050     If selgroup.iItemsCount > 0 Then
50060         'First check the VarType of Item
50070         'STRING
50080         If VarType(Item) = vbString Then
50090             For nj = 1 To selgroup.iItemsCount
50100                 If selgroup.items(nj).key = Item Then
50110                     GetItemIndex = nj
50120                     Exit Function
50130                 End If
50140             Next
50150             'When we get here, there is no Item with this key
50160             RaiseError "GetItemIndex/String: Item specified not found!"
50170             Exit Function
50180         'INTEGER
50190         ElseIf (VarType(Item) = vbInteger) Then
50200             'Does this Item Index exist?
50210             If (Item >= 1) And (Item <= selgroup.iItemsCount) Then
50220                 GetItemIndex = Item
50230                 Exit Function
50240             Else
50250                 RaiseError "GetItemIndex/Integer: Item specified not found!"
50260                 GetItemIndex = 0
50270                 Exit Function
50280             End If
50290         Else
50300             RaiseError "GetItemIndex: Item must contain String or Integer!"
50310             GetItemIndex = 0
50320             Exit Function
50330         End If
50340     'when we get here, there is no item in this group
50350     Else
50360         RaiseError "GetItemIndex: There are no Items in this group!"
50370         GetItemIndex = 0
50380         Exit Function
50390     End If
50400     'and when we get here, something else went wrong
50410     RaiseError "GetItemIndex: Unknown error!"
50420     GetItemIndex = 0
50430
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "GetItemIndex")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'**************************************************************
'Function:      GroupExists / Private
'Description:   Checks if the specified Group exists
'Parameters:    Index:      Integer, the group's index we want
'                           to check
'Result:        False:      Group dosn't exist
'               True:       Group exists
'**************************************************************

Private Function GroupExists(Index As Integer) As Boolean

    Dim dummy As String

On Error GoTo GroupError
    Select Case Index
        Case Is > 0
            dummy = cGroups(Index).key
            GroupExists = True
            Exit Function
        Case -1
            dummy = m_SpecialGroup.key
            GroupExists = True
            Exit Function
        Case -2
            dummy = m_DetailsGroup.key
            GroupExists = True
            Exit Function
        Case Else
            GroupExists = False
            Exit Function
    End Select
GroupError:
    GroupExists = False
    Err.Clear
End Function

Public Function GetSelectedGroup() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetSelectedGroup = m_SelectedGroup
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "GetSelectedGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetSelectedItem() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetSelectedItem = m_SelectedItem
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "GetSelectedItem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

'**************************************************************
'Function:      GetIconIndex / Private
'Description:   Returns the index of an image from ImageList,
'               passed as variant (Index or Key)
'               iIcon can only be String or Integer here
'Parameters:    iIcon:   Key or Index of Imagelist
'Result:        -1 for no icon found
'               Icon index for success
'**************************************************************

Private Function GetIconIndex(iIcon As Variant) As Integer

        Dim i As Integer, iLCnt As Integer
        
        On Error GoTo NoImage
        'Parameter NOT string or integer?
        If (VarType(iIcon) <> vbInteger) And (VarType(iIcon) <> vbString) Then
            RaiseError "GetIconIndex: iIcon not of required Type (String or Integer)!"
            GetIconIndex = -1
            Exit Function
        End If
        
        If iImgLType = 1 Then
            iLCnt = m_objImageList.ListImages.Count
        ElseIf iImgLType = 2 Then
            iLCnt = m_objImageList.ImageCount
        End If
        'Key was passed
        If VarType(iIcon) = vbString Then
            'get icon index
            For i = 1 To iLCnt
                If m_objImageList.ListImages(i).key = iIcon Then
                    'we did find the Icons index
                    GetIconIndex = i
                    Exit Function
                End If
            Next i
            'when we got here the string doesn't match
            RaiseError "GetIconIndex: icon with key " & iIcon & " doesn't exist!"
            GetIconIndex = -1
            Exit Function
        End If
        'Index was passed
        If iIcon >= 1 Or iIcon <= iLCnt Then
            GetIconIndex = iIcon
        Else
            RaiseWarning "GetIconIndex: invalid Image Index!"
            GetIconIndex = -1
        End If
Exit Function

NoImage:
    'No imagelist was selected
    RaiseWarning "No Defined Imagelist"
    GetIconIndex = -1
End Function

'*************************************************************
'
'   Public Functions
'
'   I'll try to add each element in runtime. I'll provide
'   all the needed functions Add groups, add items, clear,
'   and a event response for a click on each element
'
'**************************************

' Desc: Add a Group to the control
' Some parameters Still don't work, cuz I'm implementing changes.
Public Sub AddGroup(sKey As String, sCaption As String, Optional iType As Integer, Optional imgIcon As Picture, Optional imgBackground As Picture, Optional lMaskColor As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_NotOnUse = 1
50020     iGroups = iGroups + 1
50030     ReDim Preserve cGroups(iGroups)
50040     With cGroups(iGroups)
50050         .Caption = sCaption
50060         .key = sKey
50070         '.Icon = iIcon
50080         .bExpanded = True
50090     End With
50100     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "AddGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Add a Item to a group in the the control
Public Sub AddItem(Group, sKey As String, sCaption As String, Optional iIcon As Variant) 'Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020     Dim iCurrentGroup As Integer, i As Integer
50030
50040     If Not IsMissing(iIcon) Then
50050         iIcon = GetIconIndex(iIcon)
50060     Else
50070         iIcon = -1
50080     End If
50090
50100     iCurrentGroup = GetGroupsByKeyN(Group)
50110     m_NotOnUse = 1
50120
50130     If iCurrentGroup = -1 Then
50140         m_SpecialGroup.iItemsCount = m_SpecialGroup.iItemsCount + 1 'Get Current count (+1)
50150         'Debug.Print "group " & iCurrentGroup & " has " & cGroups(iCurrentGroup).iItemsCount & "Items"
50160         ReDim Preserve m_SpecialGroup.items(m_SpecialGroup.iItemsCount) 'Redim array
50170         With m_SpecialGroup.items(m_SpecialGroup.iItemsCount)
50180             .key = sKey
50190             .Caption = sCaption
50200             .sParent = "Special Group"
50210             .Index = m_SpecialGroup.iItemsCount
50220             .Icon = iIcon
50230         End With
50240     Else
50250         If iCurrentGroup = -3 Then
50260             RaiseWarning "Can't assign items to the Especified group"
50270             Exit Sub
50280         End If
50290         If iCurrentGroup = 0 Then GoTo noSuchGroup
50300         cGroups(iCurrentGroup).iItemsCount = cGroups(iCurrentGroup).iItemsCount + 1 'Get Current count (+1)
50310         ReDim Preserve cGroups(iCurrentGroup).items(cGroups(iCurrentGroup).iItemsCount) 'Redim array
50320         With cGroups(iCurrentGroup).items(cGroups(iCurrentGroup).iItemsCount)
50330             .key = sKey
50340             .Caption = sCaption
50350             .sParent = Group
50360             .Index = cGroups(iCurrentGroup).iItemsCount
50370             .Icon = iIcon
50380         End With
50390     End If
50400     UserControl_Paint
50410     Exit Sub
noSuchGroup:
50430     RaiseWarning "The group '" & Group & "' doesn't exist"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "AddItem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Set the image list object where we get the icons.
Public Sub SetImageList(ByRef ImageListObj As Object)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Set m_objImageList = ImageListObj
50020     '**********************************
50030         If TypeOf m_objImageList Is ImageList Then
50040             iImgLType = 1
50050         ElseIf TypeName(ImageListObj) = "vbalImageList" Then
50060             iImgLType = 2
50070         Else
50080             iImgLType = 0
50090             'its possible to raise an error here but not really needed?
50100         End If
50110     '**********************************
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetImageList")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Set Up the Special Group (there is only a special group in each control)
Public Sub AddSpecialGroup(Caption As String, Optional Icon As Picture, Optional background As Picture)
    m_bSpecialGroup = True
    m_SpecialGroup.Caption = Caption
    m_SpecialGroup.key = "Special Group"
    m_SpecialGroup.bExpanded = True
    m_NotOnUse = 1
    On Error Resume Next
    Set m_SpecialGroupIcon = Icon
    Set m_SpecialGroupBackground = background
    UserControl_Paint
End Sub

' Desc: Hide the special group
Public Sub HideSpecialGroup()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_bSpecialGroup = False
50020     UserControl_Paint
50030     UserControl.Refresh
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "HideSpecialGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Setup the Details Group in the control.
Public Sub AddDetailsGroup(Caption As String, sTittle As String, sDetails As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_NotOnUse = 1
50020     m_bDetailsGroup = True
50030     m_DetailsGroup.Caption = Caption
50040     m_DetailsGroup.key = "Details Group"
50050     m_DetailsGroup.Caption = Caption
50060     m_DetailsGroupTittle = sTittle
50070     m_DetailsGroupText = sDetails
50080     m_DetailsGroup.bExpanded = True
50090     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "AddDetailsGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Set the Details group Text
Public Sub SetDetailsText(sDetails As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_DetailsGroupText = sDetails
50020     m_DetailsGroup.bExpanded = True
50030     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetDetailsText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Hide the Details Group
Public Sub HideDetailsGroup()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     m_bDetailsGroup = False
50020     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "HideDetailsGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Opens a url link
Public Function OpenLink(sLink As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     OpenLink = ShellExecute(hwnd, "open", sLink, vbNull, vbNull, 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "OpenLink")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


' Desc: try to explain where the hell does all this come from.
Public Sub About()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     MsgBox "isExplorerBar Control." & vbCrLf & _
            "Developed By: Fred.cpp" & vbCrLf & _
            "HomePage: http://mx.geocities.com/fred_cpp/isexplorerar.htm" & vbCrLf & _
            "Description: this is a control that emulates almost all the functionality of the standard " & vbCrLf & _
            "Windows Explorer Bar. Uses the Windows Theme currently installed.", vbInformation, "isExplorerBar"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "About")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Clear all the structure of the Control
Public Sub ClearStructure()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Clear all the icons and groups
50020     Dim ni As Integer
50030     Dim tmpCtl
50040     Dim btmpAllowUpdates As Boolean
50050     DoEvents
50060     'Clear Special Group Items
50070     m_bSpecialGroup = False
50080     ReDim m_SpecialGroup.items(0)
50090     m_SpecialGroup.iItemsCount = 0
50100     'Clear Details Group
50110     m_bDetailsGroup = False
50120     'Clear groups
50130     'clear Childs
50140     For ni = m_pChild.LBound To m_pChild.UBound
50150         If ni <> 0 Then
50160             'm_pChild(ni).Visible = False
50170             For Each tmpCtl In UserControl.ContainedControls
50180                 If tmpCtl.Name = m_pChild(ni).Tag Then
50190                     tmpCtl.Visible = False
50200                 End If
50210             Next
50220             Unload m_pChild(ni)
50230         End If
50240     Next ni
50250     'Clear Groups
50260     ReDim cGroups(0)
50270     'Clear Counter
50280     iGroups = 0
50290     'Refresh Control
50300     btmpAllowUpdates = m_AllowRedraw
50310     UserControl.MousePointer = 0
50320     m_AllowRedraw = True
50330     UserControl_Paint
50340     UserControl.Refresh
50350     m_AllowRedraw = btmpAllowUpdates
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "ClearStructure")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Clear all the structure of the Selected group
'       if you will change lots of groups, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub ClearGroup(Group)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim iGroupIndex As Integer
50020
50030     iGroupIndex = GetGroupsByKeyN(Group)
50040
50050     'Clear all the icons in the selected group
50060     If iGroupIndex = -1 Then
50070         'clear special group Items
50080         ReDim m_SpecialGroup.items(0)
50090     Else
50100         'Clear a normal group
50110         ReDim cGroups(iGroupIndex).items(0)
50120         cGroups(iGroupIndex).iItemsCount = 0
50130     End If
50140     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "ClearGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Clear all the structure of the Selected group
'       if you will change lots of groups, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetGroupChild(Group, pChild As PictureBox, Optional pChildPointer As Integer = 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim iGroupIndex As Integer
50020
50030     iGroupIndex = GetGroupsByKeyN(Group)
50040
50050     'Setup the Item Child.
50060     If iGroupIndex = -1 Then
50070         Set m_SpecialGroup.pChild = pChild 'ReDim m_SpecialGroup.items(0)
50080         pChild.ScaleMode = 3
50090         pChild.MousePointer = pChildPointer    'set Pointer
50100     Else
50110         'Clear a normal group
50120         'ReDim cGroups(iGroupIndex).items(0)
50130         Set cGroups(iGroupIndex).pChild = pChild
50140         pChild.ScaleMode = 3
50150         pChild.MousePointer = pChildPointer     'set Pointer
50160         Load m_pChild(iGroupIndex)
50170         Set m_pChild(iGroupIndex) = cGroups(iGroupIndex).pChild
50180         m_pChild(iGroupIndex).ScaleMode = 3
50190         m_pChild(iGroupIndex).Tag = pChild.Name
50200         m_pChild(iGroupIndex).AutoRedraw = True
50210     End If
50220     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetGroupChild")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Expand an Especified group
Public Sub ExpandGroup(Group, Optional bExpand As Boolean = True)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim iGroupIndex As Integer
50020
50030     iGroupIndex = GetGroupsByKeyN(Group)
50040
50050     'Colapse the selected group
50060     If iGroupIndex = -1 Then
50070         'Colapse Special Group
50080         If IsMissing(bExpand) Then bExpand = Not m_SpecialGroup.bExpanded
50090         m_SpecialGroup.bExpanded = bExpand
50100     ElseIf iGroupIndex = -2 Then
50110         'Colapse the selected Group
50120         If IsMissing(bExpand) Then bExpand = Not m_DetailsGroup.bExpanded
50130         m_DetailsGroup.bExpanded = bExpand
50140     Else
50150         'Colapse the selected Group
50160         If IsMissing("bExpand") Then bExpand = Not cGroups(iGroupIndex).bExpanded
50170         cGroups(iGroupIndex).bExpanded = bExpand
50180     End If
50190     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "ExpandGroup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


'Desc:  This routine will change the text of a item.
'       if you will change lots of items, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetGroupCaption(Group, sNewCaption As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Set the icon of a item
50020     Dim iGroupIndex As Integer, iItemIndex As Integer
50030     Dim nj As Integer
50040
50050     iGroupIndex = GetGroupsByKeyN(Group)
50060
50070     If iGroupIndex = -3 Then
50080         Exit Sub
50090     ElseIf iGroupIndex = -2 Then
50100         m_DetailsGroup.Caption = sNewCaption
50110         UserControl_Paint
50120         Exit Sub
50130     ElseIf iGroupIndex = -1 Then
50140         m_SpecialGroup.Caption = sNewCaption
50150         UserControl_Paint
50160         Exit Sub
50170     Else
50180         cGroups(iGroupIndex).Caption = sNewCaption
50190         UserControl_Paint
50200         Exit Sub
50210     End If
50220 Exit Sub
50230     'Item not found
50240     RaiseError "The group Doesn't Exist"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetGroupCaption")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Desc:  This routine will change the icon of a item.
'       if you will change lots of items, you might
'       want to prevent redrawing using the
'       DisableUpdates method

Public Sub SetItemIcon(Group, Item, iNewIcon As Variant, Optional bUpdate As Boolean = True)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Set the icon of a item
50020     Dim iGroupIndex As Integer, iItemIndex As Integer
50030     Dim nj As Integer
50040
50050     iNewIcon = GetIconIndex(iNewIcon)
50060
50070     iGroupIndex = GetGroupsByKeyN(Group)
50080
50090     If iGroupIndex = -3 Then
50100         RaiseError "The Group '" & Group & "' doesn't exist"
50110         Exit Sub
50120     ElseIf iGroupIndex = -2 Then
50130         RaiseError "Details Group hasn't Child Items!"
50140         Exit Sub
50150     ElseIf iGroupIndex = -1 Then
50160         iItemIndex = GetItemIndex(m_SpecialGroup, Item)
50170         If iItemIndex >= 1 Then
50180             m_SpecialGroup.items(iItemIndex).Icon = iNewIcon
50190             'RedrawItem iGroupIndex, iItemIndex, 1
50200             UserControl_Paint
50210             Exit Sub
50220         End If
50230     Else
50240         If GroupExists(iGroupIndex) Then
50250             iItemIndex = GetItemIndex(cGroups(iGroupIndex), Item)
50260             If iItemIndex >= 1 Then
50270                 'We got the groupindex id and item index
50280                 cGroups(iGroupIndex).items(iItemIndex).Icon = iNewIcon
50290                 'RedrawItem iGroupIndex, iItemIndex, 1
50300                 UserControl_Paint
50310                 Exit Sub
50320             End If
50330         Else
50340             RaiseError "The Group '" & Group & "' doesn't exist"
50350             Exit Sub
50360         End If
50370     End If
50380     'When we get here, there shure was an error shown in func GetItemIndex
50390     'So we need not to raise another error here
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetItemIcon")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


'Desc:  This routine will change the text of a item.
'       if you will change lots of items, you might
'       want to prevent redrawing using the
'       DisableUpdates method
Public Sub SetItemText(Group, Item, sNewCaption As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Set the text of a item
50020     Dim iGroupIndex As Integer, iItemIndex As Integer
50030     Dim nj As Integer
50040
50050     iGroupIndex = GetGroupsByKeyN(Group)
50060
50070     If iGroupIndex = -3 Then
50080         RaiseError "The Group '" & Group & "' doesn't exist"
50090         Exit Sub
50100     ElseIf iGroupIndex = -2 Then
50110         RaiseError "Details Group hasn't Child Items!"
50120         Exit Sub
50130     ElseIf iGroupIndex = -1 Then
50140         iItemIndex = GetItemIndex(m_SpecialGroup, Item)
50150         If iItemIndex >= 1 Then
50160             m_SpecialGroup.items(iItemIndex).Caption = sNewCaption
50170             'RedrawItem iGroupIndex, iItemIndex, 1
50180             UserControl_Paint
50190             Exit Sub
50200         End If
50210     Else
50220         If GroupExists(iGroupIndex) Then
50230             iItemIndex = GetItemIndex(cGroups(iGroupIndex), Item)
50240             If iItemIndex >= 1 Then
50250                 'We got the groupindex id and item index
50260                 cGroups(iGroupIndex).items(iItemIndex).Caption = sNewCaption
50270                 'RedrawItem iGroupIndex, iItemIndex, 1
50280                 UserControl_Paint
50290                 Exit Sub
50300             End If
50310         Else
50320             RaiseError "The Group '" & Group & "' doesn't exist"
50330             Exit Sub
50340         End If
50350     End If
50360     'When we get here, there shure was an error shown in func GetItemIndex
50370     'So we need not to raise another error here
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetItemText")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

'Desc:  this function disables drawing in the control.
'       Useful if you will change the entire structure
'       and don't want to slow down the execution with
'       multiple redraws.
'       Example:
'       isExplorerBar1.DisableUdates
'       for i = 1 to List1.listcount
'           isExplorerBar1.additem "MyGroupName","Action" & i, list1.list(i)
'       next i
'       isExplorerBar1.DisableUdates False
Public Sub DisableUpdates(Optional bDisable As Boolean = True)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     'Set the internal Variable
50020     m_AllowRedraw = Not bDisable
50030     'If the control has changed, I't a good Idea update the contents
50040     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "DisableUpdates")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Description: This Sub changes the Image shown in
'       the Details group.
'       To delete the previous Image, call the routine
'       without the detailsImage Parameter.
Public Sub SetDetailsImage(Optional ByVal detailsImage As Picture)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim lmsize As Long
50020 '    lmsize = m_Width - 32
50030     Set m_DetailsPicture = detailsImage
50040     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "SetDetailsImage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Desc: Maybe you need check the Version while running
Public Function GetControlVersion() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     GetControlVersion = strCurrentVersion
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "GetControlVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Desc: As Requested, Font Property
Public Property Set Font(newFont As StdFont)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UserControl.Font.Name = newFont.Name
50020     UserControl.Font.Charset = newFont.Charset
50030     UserControl.Font.Size = newFont.Size
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Font [SET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Font() As StdFont
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Set Font = UserControl.Font
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "Font [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get UseUxThemeText() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UseUxThemeText = UxThemeText
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UseUxThemeText [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let UseUxThemeText(bNewUseUxThemeText As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     UxThemeText = bNewUseUxThemeText
50020     UserControl_Paint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "UseUxThemeText [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function EnablevbAcceleratorImagelist(bEnable As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("isExplorerBar", "EnablevbAcceleratorImagelist")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


'' Please If you found a Bug, report It. I'll fix It as soon as posible
'' If you have a suggestion or comment to this control also e-mail me
'' And please rate my work on this control
''
''  Fred.cpp
''  Last Update: 2004-7-9  / 3513 lines of code

