VERSION 5.00
Begin VB.Form frmAboutPicRT 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox pic 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   480
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   840
   End
End
Attribute VB_Name = "frmAboutPicRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Step = 310

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Me.Top = frmMain.LastAboutTop
50020  Me.Left = frmMain.LastAboutLeft + frmMain.picAbout.Width \ 2
50030  Me.Width = frmMain.picAbout.Width \ 2
50040  Me.Height = frmMain.picAbout.Height \ 2
50050  pic.Top = 0: pic.Left = 0
50060  pic.Width = Me.Width
50070  pic.Height = Me.Height
50080  With pic
50090   .Cls
50100   BitBlt .hDC, 0, 0, .Width, .Height, frmMain.picAbout.hDC, _
   frmMain.picAbout.Width \ 2 \ Screen.TwipsPerPixelX, 0, vbSrcCopy
50120  End With
50130  Me.Visible = True
50140  Timer1.Interval = 10
50150  Timer1.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAboutPicRT", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub pic_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAboutPicRT", "pic_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub pic_DblClick()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAboutPicRT", "pic_DblClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Me.Top = Me.Top + Step
50020  If Me.Top > Screen.Height Then
50030   Timer1.Enabled = False
50040   Unload Me
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAboutPicRT", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


