VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00C0FFFF&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
 ClearLogfile
 With txtLog
  .Text = ReadLogfile
  .SelStart = 0
 End With
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With
 With LanguageStrings
  Me.Caption = .LoggingLogfile
  cmdClose.Caption = .LoggingClose
  cmdClear.Caption = .LoggingClear
 End With
End Sub

Private Sub Form_Resize()
 If Me.ScaleHeight < 200 Or Me.ScaleWidth < 320 Then
  Me.Height = 200: Me.Width = 320
  Exit Sub
 End If
 With txtLog
  .Top = Me.ScaleTop
  .Left = Me.ScaleLeft
  .Width = Me.ScaleWidth
  .Height = Me.ScaleHeight - cmdClose.Height - 170
 End With
 cmdClear.Top = txtLog.Top + txtLog.Height + 150
 cmdClear.Left = txtLog.Left + 100
 cmdClose.Top = txtLog.Top + txtLog.Height + 150
 cmdClose.Left = txtLog.Left + txtLog.Width - cmdClose.Width - 100
 txtLog.Text = ReadLogfile
End Sub
