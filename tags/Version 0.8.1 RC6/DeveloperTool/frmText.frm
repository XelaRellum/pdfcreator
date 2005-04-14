VERSION 5.00
Begin VB.Form frmText 
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Close"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Filename As String

Private Sub cmd_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim fn As Long
 fn = FreeFile
 Open Filename For Input As #fn
 txt.Text = Input(LOF(fn), #fn)
 Me.Caption = Filename
 Close #fn
End Sub

Private Sub Form_Resize()
 With txt
  .Top = Me.ScaleTop + 10
  .Left = Me.ScaleLeft + 10
  .Width = Me.ScaleWidth - 20
  .Height = Me.ScaleHeight - 600
  cmd.Top = .Top + .Height + 50
  cmd.Left = .Top + .Width - cmd.Width
 End With
End Sub
