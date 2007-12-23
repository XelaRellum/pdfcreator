VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   6330
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmd 
      Caption         =   "&Down"
      Height          =   435
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Up"
      Height          =   435
      Index           =   2
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   1065
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Index           =   1
      Left            =   3870
      TabIndex        =   2
      Top             =   1275
      Width           =   1065
   End
   Begin TransTool.dmFrame dmFra 
      Height          =   975
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
      BarColorFrom    =   255
      BarColorTo      =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt 
         Appearance      =   0  '2D
         Height          =   540
         Index           =   1
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   0
         Top             =   315
         Width           =   2850
      End
   End
   Begin TransTool.dmFrame dmFra 
      Height          =   960
      Index           =   0
      Left            =   -15
      TabIndex        =   3
      Top             =   225
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   1693
      BarColorTo      =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   540
         Index           =   0
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   4
         Top             =   315
         Width           =   2850
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Save"
      Height          =   435
      Index           =   0
      Left            =   5025
      TabIndex        =   1
      Top             =   1275
      Width           =   1065
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
50030    If frmMain.lsv.SelectedItem.Text <> Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare) Then
50040     frmMain.Changed
50050     frmMain.lsv.SelectedItem.Text = Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare)
50060    End If
50070    Unload Me
50080   Case 1
50090    Unload Me
50100   Case 2
50110    If frmMain.lsv.SelectedItem.Text <> Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare) Then
50120     frmMain.Changed
50130     frmMain.lsv.SelectedItem.Text = Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare)
50140    End If
50150    frmMain.lsv.ListItems(frmMain.lsv.SelectedItem.Index - 1).Selected = True
50160    frmEdit.txt(0).Text = Replace(frmMain.lsv.SelectedItem.SubItems(5), "%n", vbCrLf)
50170    frmEdit.txt(1).Text = Replace(frmMain.lsv.SelectedItem.Text, "%n", vbCrLf)
50180    frmEdit.txt(1).SetFocus
50190    SetButtons
50200   Case 3
50210    If frmMain.lsv.SelectedItem.Text <> Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare) Then
50220     frmMain.Changed
50230     frmMain.lsv.SelectedItem.Text = Replace(txt(1).Text, vbCrLf, "%n", 1, , vbTextCompare)
50240    End If
50250    frmMain.lsv.ListItems(frmMain.lsv.SelectedItem.Index + 1).Selected = True
50260    frmEdit.txt(0).Text = Replace(frmMain.lsv.SelectedItem.SubItems(5), "%n", vbCrLf)
50270    frmEdit.txt(1).Text = Replace(frmMain.lsv.SelectedItem.Text, "%n", vbCrLf)
50280    frmEdit.txt(1).SetFocus
50290    SetButtons
50300  End Select
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

Private Sub SetButtons()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If frmMain.lsv.ListItems.Count = 1 Then
50020   cmd(2).Enabled = False
50030   cmd(3).Enabled = False
50040   Exit Sub
50050  End If
50061  Select Case frmMain.lsv.SelectedItem.Index
        Case 1
50080    cmd(2).Enabled = False
50090    cmd(3).Enabled = True
50100   Case frmMain.lsv.ListItems.Count
50110    cmd(2).Enabled = True
50120    cmd(3).Enabled = False
50130   Case Else
50140    cmd(2).Enabled = True
50150    cmd(3).Enabled = True
50160  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmEdit", "SetButtons")
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
50040  Icon = frmMain.Icon
50050  Caption = "Edit"
50060  dmFra(0).Caption = "Template text"
50070  dmFra(1).Caption = "Translated text"
50080  Set txt(1).Font = frmMain.lsv.Font
50090  ShowAcceleratorsInForm Me, True
50100  SetButtons
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
50080   With dmFra(0)
50090    .Top = 100
50100    .Left = 100
50110    .Height = (ScaleHeight - cmd(0).Height - 400) / 2
50120    .Width = ScaleWidth - 2 * .Left
50130   End With
50140   With dmFra(1)
50150    .Top = dmFra(0).Top + dmFra(0).Height + 100
50160    .Left = dmFra(0).Left
50170    .Height = dmFra(0).Height
50180    .Width = dmFra(0).Width
50190   End With
50200   With txt(0)
50210    .Top = dmFra(0).BarHeight * Screen.TwipsPerPixelY + 50
50220    .Left = 50
50230    .Height = dmFra(0).Height - .Top - 100
50240    .Width = dmFra(0).Width - 2 * .Left
50250   End With
50260   With txt(1)
50270    .Top = dmFra(1).BarHeight * Screen.TwipsPerPixelY + 50
50280    .Left = 50
50290    .Height = dmFra(1).Height - .Top - 100
50300    .Width = dmFra(1).Width - 2 * .Left
50310   End With
50320   With cmd(0)
50330    .Top = dmFra(1).Top + dmFra(1).Height + 100
50340    .Left = dmFra(1).Left + dmFra(1).Width - .Width
50350   End With
50360   With cmd(1)
50370    .Top = cmd(0).Top
50380    .Left = dmFra(1).Left
50390   End With
50400   With cmd(2)
50410    .Top = cmd(0).Top
50420    .Left = Me.Width / 2 - .Width - 50
50430   End With
50440   With cmd(3)
50450    .Top = cmd(0).Top
50460    .Left = Me.Width / 2 + 50
50470   End With
50480  End If
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
