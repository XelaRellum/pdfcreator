VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Enter Passwords"
   ClientHeight    =   2790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOwnerPass 
      Caption         =   "Owner Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
      Begin VB.TextBox txtOwnerPassRepeat 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   680
         Width           =   1695
      End
      Begin VB.TextBox txtOwnerPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblOwnerPassRepeat 
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblOwnerPass 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame fraUserPass 
      Caption         =   "User Password"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtUserPassRepeat 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   660
         Width           =   1695
      End
      Begin VB.TextBox txtUserPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblUserPassRepeat 
         Caption         =   "Repeat:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblUserPass 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bSuccess As Boolean
Public bFinished As Boolean
Public iPasswords As Integer

Private Sub CancelButton_Click()
bSuccess = False
bFinished = True
End Sub

Private Sub Form_Load()
Me.Caption = LanguageStrings.OptionsPDFEnterPasswords
fraUserPass.Caption = LanguageStrings.OptionsUserPass
fraOwnerPass.Caption = LanguageStrings.OptionsOwnerPass
lblUserPass.Caption = LanguageStrings.OptionsPDFSetPassword
lblUserPassRepeat.Caption = LanguageStrings.OptionsPDFRepeatPassword
lblOwnerPass.Caption = LanguageStrings.OptionsPDFSetPassword
lblOwnerPassRepeat.Caption = LanguageStrings.OptionsPDFRepeatPassword
OKButton.Caption = LanguageStrings.OptionsPassOK
CancelButton.Caption = LanguageStrings.OptionsPassCancel
fraUserPass.Enabled = Options.PDFUserPass
fraOwnerPass.Enabled = Options.PDFOwnerPass
End Sub

Private Sub OKButton_Click()
If txtUserPass.Text <> txtUserPassRepeat.Text Then
  MsgBox "User Passes do not match", vbCritical
  Exit Sub
End If

If txtOwnerPass.Text <> txtOwnerPassRepeat.Text Then
  MsgBox "Owner Passes do not match", vbCritical
  Exit Sub
End If

bSuccess = True
bFinished = True
End Sub


