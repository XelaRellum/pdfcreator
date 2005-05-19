VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnimation 
   Caption         =   "PDFCreator"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2100
   Icon            =   "frmAnimation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   2100
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1155
      Top             =   105
   End
   Begin MSComCtl2.Animation anmProcess 
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackStyle       =   1
      FullWidth       =   64
      FullHeight      =   64
   End
End
Attribute VB_Name = "frmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ShowAnimation(Show As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DoEvents
50020  If Show = True Then
50030    anmProcess.Visible = Show
50040    ResAnimate anmProcess, ranOpen, 100
50050    ResAnimate anmProcess, ranPlay
50060   Else
50070    ResAnimate anmProcess, ranStop
50080    ResAnimate anmProcess, ranClose
50090    Me.Height = 2520
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAnimation", "ShowAnimation")
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
50010  Dim tL As Long, BorderWidth As Long
50020  With anmProcess
50030    .Top = 0
50040    .Left = 0
50050    .Width = 260 * Screen.TwipsPerPixelX
50060    .Height = 66 * Screen.TwipsPerPixelY
50070   End With
50080    BorderWidth = 3
50090    anmProcess.Left = BorderWidth * Screen.TwipsPerPixelX
50100    anmProcess.Top = BorderWidth * Screen.TwipsPerPixelY
50110    Height = anmProcess.Height + 380 + 2 * BorderWidth * Screen.TwipsPerPixelY + 10
50120    Width = anmProcess.Width + 4 * BorderWidth * Screen.TwipsPerPixelX + 10
50130    BorderStyle = vbBSNone
50140    Caption = Caption
50150    tL = Width
50160    Width = tL - Screen.TwipsPerPixelX
50170    Width = tL
50180    DrawBorder3D Me, 4, BorderWidth
50190    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
50200  SetTopMost Me, True, True
50210 ' SetTopMost Me, False, True
50220  SetActiveWindow hwnd
50230  ShowAnimation True
50240  Timer1.Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAnimation", "Form_Load")
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
50010  If ShowAnimationWindow = False Then
50020   Timer1.Enabled = False
50030   ShowAnimation False
50040   Unload Me
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAnimation", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
