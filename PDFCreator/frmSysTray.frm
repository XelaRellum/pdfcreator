VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Sys Tray Interface"
   ClientHeight    =   1920
   ClientLeft      =   5610
   ClientTop       =   3360
   ClientWidth     =   4680
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4680
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Begin VB.Menu mnuSysTray 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 03/03/2003
' * Added Unicode support
' * Added support for new tray version (ME,2000 or above required)
' * Added support for balloon tips (ME,2000 or above required)

' frmSysTray.
' Steve McMahon
' Original version based on code supplied from Ben Baird:

'Author:
'        Ben Baird <psyborg@cyberhighway.com>
'        Copyright (c) 1997, Ben Baird
'
'Purpose:
'        Demonstrates setting an icon in the taskbar's
'        system tray without the overhead of subclassing
'        to receive events.

Private nfIconDataA As NOTIFYICONDATAA
Private nfIconDataW As NOTIFYICONDATAW

Private Const NOTIFYICONDATAA_V1_SIZE_A = 88
Private Const NOTIFYICONDATAA_V1_SIZE_U = 152
Private Const NOTIFYICONDATAA_V2_SIZE_A = 488
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
Public Event SysTrayMouseMove()
Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeOut()
Public Event BalloonClicked()

Public Enum EBalloonIconTypes
 NIIF_NONE = 0
 NIIF_INFO = 1
 NIIF_WARNING = 2
 NIIF_ERROR = 3
 NIIF_NOSOUND = &H10
End Enum

Private m_bAddedMenuItem As Boolean
Private m_iDefaultIndex As Long

Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Public Sub ShowBalloonTip(ByVal sMessage As String, Optional ByVal sTitle As String, _
 Optional ByVal eIcon As EBalloonIconTypes, Optional ByVal lTimeOutMs = 30000)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lr As Long
50020  If m_bSupportsNewVersion Then
50030    If m_bUseUnicode Then
50040      stringToArray sMessage, nfIconDataW.szInfo, 512
50050      stringToArray sTitle, nfIconDataW.szInfoTitle, 128
50060      nfIconDataW.uTimeOutOrVersion = lTimeOutMs
50070      nfIconDataW.dwInfoFlags = eIcon
50080      nfIconDataW.uFlags = NIF_INFO
50090      lr = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
50100     Else
50110      nfIconDataA.szInfo = sMessage
50120      nfIconDataA.szInfoTitle = sTitle
50130      nfIconDataA.uTimeOutOrVersion = lTimeOutMs
50140      nfIconDataA.dwInfoFlags = eIcon
50150      nfIconDataA.uFlags = NIF_INFO
50160      lr = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
50170    End If
50180   Else
50190    ' can't do it, fail silently.
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "ShowBalloonTip")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get ToolTip() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sTip As String, iPos As Long
50020  sTip = nfIconDataA.szTip
50030  iPos = InStr(sTip, Chr$(0))
50040  If (iPos <> 0) Then
50050   sTip = Left$(sTip, iPos - 1)
50060  End If
50070  ToolTip = sTip
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "ToolTip [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let ToolTip(ByVal sTip As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If (m_bUseUnicode) Then
50020    stringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
50030    nfIconDataW.uFlags = NIF_TIP
50040    Shell_NotifyIconW NIM_MODIFY, nfIconDataW
50050   Else
50060    If (sTip & Chr$(0) <> nfIconDataA.szTip) Then
50070     nfIconDataA.szTip = sTip & Chr$(0)
50080     nfIconDataA.uFlags = NIF_TIP
50090     Shell_NotifyIconA NIM_MODIFY, nfIconDataA
50100    End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "ToolTip [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get IconHandle() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  IconHandle = nfIconDataA.hIcon
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "IconHandle [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let IconHandle(ByVal hIcon As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If m_bUseUnicode Then
50020 '   If hIcon <> nfIconDataW.hIcon Then
50030     nfIconDataW.hIcon = hIcon
50040     nfIconDataW.uFlags = NIF_ICON
50050     Shell_NotifyIconW NIM_MODIFY, nfIconDataW
50060 '   End If
50070   Else
50080 '   If hIcon <> nfIconDataA.hIcon Then
50090     nfIconDataA.hIcon = hIcon
50100     nfIconDataA.uFlags = NIF_ICON
50110     Shell_NotifyIconA NIM_MODIFY, nfIconDataA
50120 '   End If
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "IconHandle [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function AddMenuItem(ByVal sCaption As String, Optional ByVal sKey As String = "", Optional ByVal bDefault As Boolean = False) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim iIndex As Long
50020  If Not m_bAddedMenuItem Then
50030    iIndex = 0
50040    m_bAddedMenuItem = True
50050   Else
50060    iIndex = mnuSysTray.UBound + 1
50070    Load mnuSysTray(iIndex)
50080  End If
50090  mnuSysTray(iIndex).Visible = True
50100  mnuSysTray(iIndex).Tag = sKey
50110  mnuSysTray(iIndex).Caption = sCaption
50120  If bDefault Then
50130   m_iDefaultIndex = iIndex
50140  End If
50150  AddMenuItem = iIndex
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "AddMenuItem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function ValidIndex(ByVal lIndex As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ValidIndex = (lIndex >= mnuSysTray.LBound And lIndex <= mnuSysTray.UBound)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "ValidIndex")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub EnableMenuItem(ByVal lIndex As Long, ByVal bState As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If ValidIndex(lIndex) Then
50020   mnuSysTray(lIndex).Enabled = bState
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "EnableMenuItem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function RemoveMenuItem(ByVal iIndex As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If ValidIndex(iIndex) Then
50030   If (iIndex = 0) Then
50040     mnuSysTray(0).Caption = ""
50050    Else
50060     ' remove the item:
50070     For i = iIndex + 1 To mnuSysTray.UBound
50080      mnuSysTray(iIndex - 1).Caption = mnuSysTray(iIndex).Caption
50090      mnuSysTray(iIndex - 1).Tag = mnuSysTray(iIndex).Tag
50100     Next i
50110     Unload mnuSysTray(mnuSysTray.UBound)
50120   End If
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "RemoveMenuItem")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Property Get DefaultMenuIndex() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DefaultMenuIndex = m_iDefaultIndex
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "DefaultMenuIndex [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let DefaultMenuIndex(ByVal lIndex As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If ValidIndex(lIndex) Then
50020    m_iDefaultIndex = lIndex
50030   Else
50040    m_iDefaultIndex = 0
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "DefaultMenuIndex [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function ShowMenu()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetForegroundWindow Me.hwnd
50020  If m_iDefaultIndex > -1 Then
50030    Me.PopupMenu mnuPopup, 0, , , mnuSysTray(m_iDefaultIndex)
50040   Else
50050    Me.PopupMenu mnuPopup, 0
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "ShowMenu")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lMajor As Long, lMinor As Long, bIsNt As Long, lr As Long
50020  GetWindowsVersion lMajor, lMinor, , , bIsNt
50030  If bIsNt Then
50040    m_bUseUnicode = True
50050    If (lMajor >= 5) Then
50060     ' 2000 or XP
50070     m_bSupportsNewVersion = True
50080    End If
50090   ElseIf (lMajor = 4) And (lMinor = 90) Then
50100    ' Windows ME
50110    m_bSupportsNewVersion = True
50120  End If
50130    'Add the icon to the system tray...
50140  If m_bUseUnicode Then
50150    With nfIconDataW
50160     .hwnd = Me.hwnd
50170     .uID = Me.Icon
50180     .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
50190     .uCallbackMessage = WM_MOUSEMOVE
50200     .hIcon = Me.Icon.handle
50210     stringToArray App.FileDescription, .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
50220     If m_bSupportsNewVersion Then
50230      .uTimeOutOrVersion = NOTIFYICON_VERSION
50240     End If
50250     .cbSize = nfStructureSize
50260    End With
50270    lr = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
50280    If m_bSupportsNewVersion Then
50290     Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
50300    End If
50310   Else
50320    With nfIconDataA
50330     .hwnd = Me.hwnd
50340     .uID = Me.Icon
50350     .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
50360     .uCallbackMessage = WM_MOUSEMOVE
50370     .hIcon = Me.Icon.handle
50380     .szTip = App.FileDescription & Chr$(0)
50390     If m_bSupportsNewVersion Then
50400      .uTimeOutOrVersion = NOTIFYICON_VERSION
50410     End If
50420     .cbSize = nfStructureSize
50430    End With
50440    lr = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
50450    If m_bSupportsNewVersion Then
50460     lr = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
50470    End If
50480  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub stringToArray(ByVal sString As String, bArray() As Byte, _
 ByVal lMaxSize As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim B() As Byte, i As Long, j As Long
50020  If Len(sString) > 0 Then
50030   B = sString
50040   For i = LBound(B) To UBound(B)
50050    bArray(i) = B(i)
50060    If (i = (lMaxSize - 2)) Then
50070     Exit For
50080    End If
50090   Next i
50100   For j = i To lMaxSize - 1
50110    bArray(j) = 0
50120   Next j
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "stringToArray")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
Private Function unicodeSize(ByVal lSize As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If m_bUseUnicode Then
50020    unicodeSize = lSize * 2
50030   Else
50040    unicodeSize = lSize
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "unicodeSize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Property Get nfStructureSize() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If m_bSupportsNewVersion Then
50020    If m_bUseUnicode Then
50030      nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
50040     Else
50050      nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
50060    End If
50070   Else
50080    If m_bUseUnicode Then
50090      nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
50100     Else
50110      nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
50120    End If
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "nfStructureSize [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lx As Long
50020  ' VB manipulates the x value according to scale mode:
50030  ' we must remove this before we can interpret the
50040  ' message windows was trying to send to us:
50050  lx = ScaleX(x, Me.ScaleMode, vbPixels)
50061  Select Case lx
        Case WM_MOUSEMOVE
50080    RaiseEvent SysTrayMouseMove
50090   Case WM_LBUTTONUP
50100    RaiseEvent SysTrayMouseDown(vbLeftButton)
50110   Case WM_LBUTTONUP
50120    RaiseEvent SysTrayMouseUp(vbLeftButton)
50130   Case WM_LBUTTONDBLCLK
50140    RaiseEvent SysTrayDoubleClick(vbLeftButton)
50150   Case WM_RBUTTONDOWN
50160    RaiseEvent SysTrayMouseDown(vbRightButton)
50170   Case WM_RBUTTONUP
50180    RaiseEvent SysTrayMouseUp(vbRightButton)
50190   Case WM_RBUTTONDBLCLK
50200    RaiseEvent SysTrayDoubleClick(vbRightButton)
50210   Case NIN_BALLOONSHOW
50220    RaiseEvent BalloonShow
50230   Case NIN_BALLOONHIDE
50240    RaiseEvent BalloonHide
50250   Case NIN_BALLOONTIMEOUT
50260    RaiseEvent BalloonTimeOut
50270   Case NIN_BALLOONUSERCLICK
50280    RaiseEvent BalloonClicked
50290   End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "Form_MouseMove")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If m_bUseUnicode Then
50020    Shell_NotifyIconW NIM_DELETE, nfIconDataW
50030   Else
50040    Shell_NotifyIconA NIM_DELETE, nfIconDataA
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "Form_QueryUnload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnuSysTray_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  RaiseEvent MenuClick(Index, mnuSysTray(Index).Tag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "mnuSysTray_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, _
 Optional ByRef lMinor = 0, Optional ByRef lRevision = 0, _
 Optional ByRef lBuildNumber = 0, Optional ByRef bIsNt = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lr As Long
50020  lr = GetVersion()
50030  lBuildNumber = (lr And &H7F000000) \ &H1000000
50040  If (lr And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
50050  lRevision = (lr And &HFF0000) \ &H10000
50060  lMinor = (lr And &HFF00&) \ &H100
50070  lMajor = (lr And &HFF)
50080  bIsNt = ((lr And &H80000000) = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmSysTray", "GetWindowsVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
