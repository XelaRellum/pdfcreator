VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4395
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fra 
      Caption         =   "Frame1"
      Height          =   960
      Index           =   1
      Left            =   2205
      TabIndex        =   5
      Top             =   105
      Width           =   2115
      Begin VB.TextBox txt 
         Height          =   540
         Index           =   1
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   0
         Top             =   210
         Width           =   1905
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Frame1"
      Height          =   960
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   105
      Width           =   2115
      Begin VB.TextBox txt 
         BackColor       =   &H8000000F&
         Height          =   540
         Index           =   0
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   3
         Top             =   210
         Width           =   1905
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   435
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1155
      Width           =   1170
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Save"
      Height          =   435
      Index           =   0
      Left            =   1260
      TabIndex        =   1
      Top             =   1155
      Width           =   1170
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0
50030    If frmMain.lsv.SelectedItem.Index > 0 Then
50040     frmMain.lsv.SelectedItem.Text = Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare)
50050    End If
50060  End Select
50070  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmEdit", "cmd_Click")
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
50010  Height = 0.6 * Screen.Height
50020  Width = 0.6 * Screen.Width
50030  Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
50040  With frmMain
50050   Icon = .Icon
50060   Caption = "Edit"
50070  End With
50080  fra(0).Caption = "Template text"
50090  fra(1).Caption = "Translated text"
50100  With txt(1)
50110   Set .Font = frmMain.lsv.Font
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmEdit", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If WindowState <> vbMinimized Then
50020   If Width < 3000 Then
50030    Width = 3000
50040   End If
50050   If Height < 3000 Then
50060    Height = 3000
50070   End If
50080   With fra(0)
50090    .Top = 0
50100    .Left = 0
50110    .Height = (ScaleHeight - cmd(0).Height) / 2 - 10
50120    .Width = ScaleWidth - .Left
50130   End With
50140   With fra(1)
50150    .Top = fra(0).Top + fra(0).Height
50160    .Left = fra(0).Left
50170    .Height = fra(0).Height
50180    .Width = fra(0).Width
50190   End With
50200   With txt(0)
50210    .Top = 250
50220    .Left = 50
50230    .Height = fra(0).Height - .Top - 100
50240    .Width = fra(0).Width - 2 * .Left
50250   End With
50260   With txt(1)
50270    .Top = 250
50280    .Left = 50
50290    .Height = fra(1).Height - .Top - 100
50300    .Width = fra(1).Width - 2 * .Left
50310   End With
50320   With cmd(0)
50330    .Top = fra(1).Top + fra(1).Height + 20
50340    .Left = fra(1).Left + fra(1).Width - .Width
50350   End With
50360   With cmd(1)
50370    .Top = cmd(0).Top
50380    .Left = fra(1).Left
50390   End With
50400  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmEdit", "Form_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
