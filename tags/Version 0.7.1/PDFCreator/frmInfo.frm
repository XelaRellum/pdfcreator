VERSION 5.00
Begin VB.Form frmInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   Picture         =   "frmInfo.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picTitle 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   2160
      Picture         =   "frmInfo.frx":75342
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.PictureBox picPDF 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   345
      Picture         =   "frmInfo.frx":7A79C
      ScaleHeight     =   1035
      ScaleWidth      =   2085
      TabIndex        =   2
      Top             =   990
      Width           =   2085
   End
   Begin VB.PictureBox picCredits 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   1500
      ScaleHeight     =   4575
      ScaleWidth      =   4350
      TabIndex        =   3
      Top             =   1260
      Width           =   4350
   End
   Begin VB.PictureBox picForeground 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   16000
      Left            =   -945
      Picture         =   "frmInfo.frx":81912
      ScaleHeight     =   16005
      ScaleWidth      =   4350
      TabIndex        =   0
      Top             =   -10710
      Visible         =   0   'False
      Width           =   4350
   End
   Begin VB.Image imgClose 
      Height          =   600
      Left            =   1455
      Picture         =   "frmInfo.frx":1260DC
      Top             =   105
      Width           =   600
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private m_blnUnloading As Boolean

Private Sub AnimateScroller()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lngSecondCurrX As Long, lngCurrentX As Long, intCounter As Integer
50020
50030  While Not m_blnUnloading
50040   lngCurrentX = lngCurrentX + 1
50050   If intCounter = 40 Then intCounter = -1
50060   intCounter = intCounter + 1
50070   picCredits.Cls
50080   If lngCurrentX + picCredits.ScaleHeight > picForeground.ScaleHeight Then
50090     lngSecondCurrX = picForeground.ScaleHeight - lngCurrentX
50100    Else
50110     lngSecondCurrX = picCredits.ScaleHeight
50120   End If
50130
50140   Call TransBlt(picCredits.hDC, 0, 0, picForeground.ScaleWidth, lngSecondCurrX, picForeground.hDC, 0, lngCurrentX, vbBlack)
50150
50160   If lngSecondCurrX < picCredits.ScaleHeight Then
50170    Call TransBlt(picCredits.hDC, 0, lngSecondCurrX, picForeground.ScaleWidth, picCredits.ScaleHeight - lngSecondCurrX, picForeground.hDC, 0, 0, vbBlack)
50180    'DoEvents
50190   End If
50200
50210   If lngSecondCurrX = 0 Then lngCurrentX = 0
50220   DoEvents
50230  Wend
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "AnimateScroller")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Activate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_blnUnloading = False
50020  Call AnimateScroller
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "Form_Activate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim rctPicRect As RECT, Version As String
50020
50030  Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
50040  Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
50050  MakeFormTransparent Me, vbMagenta
50060
50070  picCredits.ScaleMode = vbPixels
50080  picCredits.AutoRedraw = True
50090
50100  picForeground.ScaleMode = vbPixels
50110  picForeground.AutoRedraw = True
50120
50130  Version = GetProgramRelease
50140
50150  With picTitle
50160   .ForeColor = RGB(7, 16, 127)
50170   .CurrentX = 15
50180   .CurrentY = 35
50190   picTitle.Print Version
50200  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ReleaseCapture
50020  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "Form_MouseDown")
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
50010  m_blnUnloading = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "Form_QueryUnload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub imgClose_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "imgClose_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub picCredits_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ReleaseCapture
50020  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "picCredits_MouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub picForeground_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ReleaseCapture
50020  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "picForeground_MouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub picPDF_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ReleaseCapture
50020  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "picPDF_MouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ReleaseCapture
50020  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmInfo", "picTitle_MouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub