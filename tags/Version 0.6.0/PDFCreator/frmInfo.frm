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
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   2160
      Picture         =   "frmInfo.frx":75342
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   3
      Top             =   150
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
      TabIndex        =   1
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
      Picture         =   "frmInfo.frx":14929C
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
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private m_blnUnloading As Boolean

Private Sub AnimateScroller()
 Dim lngSecondCurrX As Long, lngCurrentX As Long, intCounter As Integer
 
 While Not m_blnUnloading
  lngCurrentX = lngCurrentX + 1
  If intCounter = 40 Then intCounter = -1
  intCounter = intCounter + 1
  picCredits.Cls
  If lngCurrentX + picCredits.ScaleHeight > picForeground.ScaleHeight Then
    lngSecondCurrX = picForeground.ScaleHeight - lngCurrentX
   Else
    lngSecondCurrX = picCredits.ScaleHeight
  End If
   
  Call TransBlt(picCredits.hDC, 0, 0, picForeground.ScaleWidth, lngSecondCurrX, picForeground.hDC, 0, lngCurrentX, vbBlack)
      
  If lngSecondCurrX < picCredits.ScaleHeight Then
   Call TransBlt(picCredits.hDC, 0, lngSecondCurrX, picForeground.ScaleWidth, picCredits.ScaleHeight - lngSecondCurrX, picForeground.hDC, 0, 0, vbBlack)
   'DoEvents
  End If
      
  If lngSecondCurrX = 0 Then lngCurrentX = 0
  DoEvents
 Wend
End Sub

Private Sub Form_Activate()
 m_blnUnloading = False
 Call AnimateScroller
End Sub

Private Sub Form_Load()
 Dim rctPicRect As RECT

 Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips)
 Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips)
 MakeFormTransparent Me, vbMagenta
    
 picCredits.ScaleMode = vbPixels
 picCredits.AutoRedraw = True
    
 picForeground.ScaleMode = vbPixels
 picForeground.AutoRedraw = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 m_blnUnloading = True
End Sub

Private Sub imgClose_Click()
 Unload Me
End Sub

Private Sub picCredits_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub picForeground_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub picPDF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 ReleaseCapture
 SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
