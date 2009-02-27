VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
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
50020  Me.Icon = LoadResPicture(2120, vbResIcon)
50030  With anmProcess
50040    .Top = 0
50050    .Left = 0
50060    .Width = 260 * Screen.TwipsPerPixelX
50070    .Height = 66 * Screen.TwipsPerPixelY
50080   End With
50090    BorderWidth = 3
50100    anmProcess.Left = BorderWidth * Screen.TwipsPerPixelX
50110    anmProcess.Top = BorderWidth * Screen.TwipsPerPixelY
50120    Height = anmProcess.Height + 380 + 2 * BorderWidth * Screen.TwipsPerPixelY + 10
50130    Width = anmProcess.Width + 4 * BorderWidth * Screen.TwipsPerPixelX + 10
50140    BorderStyle = vbBSNone
50150    Caption = Caption
50160    tL = Width
50170    Width = tL - Screen.TwipsPerPixelX
50180    Width = tL
50190    DrawBorder3D Me, 4, BorderWidth
50200    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
50210  SetTopMost Me, True, True
50220 ' SetTopMost Me, False, True
50230  SetActiveWindow hwnd
50240  ShowAnimation True
50250  Timer1.Enabled = True
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

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  'Dummy
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmAnimation", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

